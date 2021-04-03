VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'调用模块编号
Public Enum PASS_MODEL
    PM_门诊编辑 = 0
    PM_住院编辑 = 1
    PM_住院医嘱清单 = 2
    PM_护士校对 = 3
    PM_门诊医嘱清单 = 4
    PM_住院首页 = 5
    PM_处方发药 = 1341        '1341    药品处方发药
    PM_部门发药 = 1342        '1342    药品部门发药
    PM_PIVA管理 = 1345        '1345    PIVA管理
    PM_门诊处方审查 = 1351     '1351:     门诊处方审查
    PM_住院药嘱审查 = 1352     '1352: 住院药嘱审查
End Enum

'接口类型枚举
Public Enum G_PASS_TYPE
    UNPASS = 0          '未启用
    MK = 1              '美康
    DT = 2              '大通
    TYT = 3             '太元通
    YWS = 4             '药卫士
    HZYY = 5            '杭州逸曜
    ZL = 6              '中联合理用药
End Enum

'保持属性值的局部变量
Private mvarPassType As G_PASS_TYPE  '接口类型
Private mvarPassVersion As String    '版本号
Private mvarPassModel As Integer     '0-缺省方式;1-参数设置  100315

Private Property Let PassModel(ByVal vData As Integer)
    mvarPassModel = vData
End Property

Public Property Get PassModel() As Integer
    PassModel = mvarPassModel
End Property

Private Property Let PassType(ByVal vData As G_PASS_TYPE)
    mvarPassType = vData
End Property

Public Property Get PassType() As G_PASS_TYPE
    PassType = mvarPassType
End Property

Private Property Let PassVersion(ByVal vData As String)
    mvarPassVersion = vData
End Property

Public Property Get PassVersion() As String
    PassVersion = mvarPassVersion
End Property

Public Function HaveRecipNo() As Boolean
'功能:是否允许生成处方号
    HaveRecipNo = gbytPass = HZYY Or (gbytPass = MK And gstrVersion = "4.0")
End Function

Public Function zlPassInit(cnMain As ADODB.Connection, ByVal lngSys As Long, Optional ByVal lngModel As Long) As Boolean
'功能:
'参数:lngModel-可选 PM_住院首页 （太元通过敏源接口）

    glngSys = lngSys
    glngModel = lngModel
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    Set gcnOracle = cnMain
    '权限判断
    If Not PassCheckPrivs(lngModel, True) Then Exit Function
    '初始化
    If Not InitSysPar Then Exit Function
    Call GetUserInfo
    
    '设置PASS类型
    PassType = gbytPass
    PassVersion = gstrVersion
    
    If (gbytPass = MK And gstrVersion = "3.0") Then
        gblnInitOK = False '具体执行接口时初始化
    Else
        gblnInitOK = True
        If Not PassInitialize Then
            gbytPass = UNPASS
            PassType = UNPASS
            Exit Function
        End If
    End If
    zlPassInit = True
End Function

Public Sub zlPassCommandBarExe(ByRef objMap As clsPassMap, ByVal lngIndex As Long, Optional ByRef blnNoSave As Boolean)
'功能:住院医嘱清单管理合理用药监测执行
'参数:
'      lngIndex-菜单对象索引
'      gint场合- '调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    Dim strDrugCode As String
    Dim strDrugName As String
    Dim strQueryType As String
    Dim strFuncName As String
    Dim lngX As Long, lngY As Long
    Dim lngReturn As Long
    Dim objPoint As RECT
    
    On Error GoTo errH

    If Not PassMap(objMap) Then Exit Sub
    
    With gobjAdvice
        If gbytPass = MK Then    '美康审查
            If gstrVersion = "3.0" Then
                Select Case lngIndex
                Case MK_IX_药物临床信息参考 To MK_IX_检验值, MK_IX_药物相互作用 To MK_IX_系统设置
                    Call PassDoCommand(gcolPASSExe("_" & lngIndex))
                Case MK_IX_用药研究
                    '用药研究
                    If glngModel = PM_住院医嘱清单 Or glngModel = PM_住院编辑 Then
                        Call InAdviceCheckWarn_MK(gcolPASSExe("_" & lngIndex))
                    ElseIf glngModel = PM_门诊编辑 Or glngModel = PM_门诊医嘱清单 Then
                        Call OutAdviceCheckWarn_MK(gcolPASSExe("_" & lngIndex))
                    End If
                Case MK_IX_警告
                    '警告:有警示值(不为空),且大于0-蓝灯
                    If glngModel = PM_住院医嘱清单 Or glngModel = PM_住院编辑 Then
                        Call InAdviceCheckWarn_MK(gcolPASSExe("_" & lngIndex), .Row)
                    ElseIf glngModel = PM_门诊编辑 Or glngModel = PM_门诊医嘱清单 Then
                        Call OutAdviceCheckWarn_MK(gcolPASSExe("_" & lngIndex), .Row)
                    End If
                    '审查
                Case MK_IX_审查
                    If glngModel = PM_住院医嘱清单 Then
                        Call InAdviceCheckWarn_MK(gcolPASSExe("_" & lngIndex))
                    ElseIf glngModel = PM_住院编辑 Then
                        Call InAdviceCheckWarn_MK(gcolPASSExe("_" & lngIndex), , , blnNoSave)
                    ElseIf glngModel = PM_门诊编辑 Or glngModel = PM_门诊医嘱清单 Then
                        Call OutAdviceCheckWarn_MK(gcolPASSExe("_" & lngIndex), , blnNoSave)
                    End If
                   
                End Select
            ElseIf gstrVersion = "4.0" Then
                '美康4.0
                Select Case glngModel
                Case PM_住院医嘱清单
                    If gint场合 = US_InDoctor Then
                        Call InAdviceCheckWarn_MK4(1, 0)  '手动审查不采集
                    End If
                Case PM_门诊医嘱清单
                    If gint场合 = US_InDoctor Then
                        Call OutAdviceCheckWarn_MK4(1, 0)  '手动审查不采集
                    End If
                Case PM_门诊编辑
                    Call OutAdviceCheckWarn_MK4(1, 0)
                Case PM_住院编辑
                    Call InAdviceCheckWarn_MK4(1, 0) '手动审查不采集
                End Select
            End If
        ElseIf gbytPass = DT Then
            If gstrVersion = "3.0" Then
                Select Case glngModel
                Case PM_住院医嘱清单
                    If gint场合 = US_InDoctor Then
                        Call InAdviceCheckWarn_DT
                    End If
                Case PM_门诊医嘱清单
                    If gint场合 = US_InDoctor Then
                        Call OutAdviceCheckWarn_DT
                    End If
                End Select
            ElseIf gstrVersion = "4.0" Then
                Select Case glngModel
                Case PM_住院医嘱清单
                    If gint场合 = US_InDoctor Then
                        Call AdviceCheckWarn_DTBS(1)
                    End If
                Case PM_门诊医嘱清单
                    If gint场合 = US_InDoctor Then
                        Call AdviceCheckWarn_DTBS(1)
                    End If
                End Select
            End If
        ElseIf gbytPass = YWS Then
            Select Case glngModel
            Case PM_住院医嘱清单
                If gint场合 = US_InDoctor Then
                    Call InAdviceCheckWarn_YWS
                End If
            Case PM_门诊医嘱清单
                If gint场合 = US_InDoctor Then
                    Call OutAdviceCheckWarn_YWS
                End If
            End Select
        ElseIf gbytPass = TYT Then   '太元通药嘱审查
            If glngModel = PM_住院医嘱清单 Or glngModel = PM_住院编辑 Then
                Select Case lngIndex
                Case TYT_用药规范
                    Call InAdviceCheckWarn_TYT(TYT_用药规范)
                Case TYT_药嘱审查
                    If glngModel = PM_住院医嘱清单 Then
                        Call InAdviceCheckWarn_TYT(TYT_药嘱审查)
                    Else
                        Call InAdviceCheckWarn_TYT(TYT_药嘱审查, , , blnNoSave)
                    End If
                Case TYT_药品提示
                    Call InAdviceCheckWarn_TYT(TYT_药品提示, .Row)
                Case TYT_医药知识库
                    Call InAdviceCheckWarn_TYT(TYT_医药知识库)
                Case TYT_系统配置
                    Call InAdviceCheckWarn_TYT(TYT_系统配置)
                End Select
            ElseIf glngModel = PM_门诊编辑 Or glngModel = PM_门诊医嘱清单 Then
                Select Case lngIndex
                Case TYT_用药规范
                    Call OutAdviceCheckWarn_TYT(TYT_用药规范)
                Case TYT_药嘱审查
                    If glngModel = PM_门诊医嘱清单 Then
                        Call OutAdviceCheckWarn_TYT(TYT_药嘱审查)
                    Else
                        Call OutAdviceCheckWarn_TYT(TYT_药嘱审查, , blnNoSave)
                    End If
                Case TYT_药品提示
                    Call OutAdviceCheckWarn_TYT(TYT_药品提示, .Row)
                Case TYT_医药知识库
                    Call OutAdviceCheckWarn_TYT(TYT_医药知识库)
                Case TYT_系统配置
                    Call OutAdviceCheckWarn_TYT(TYT_系统配置)
                End Select
            End If
        ElseIf gbytPass = HZYY Then   '杭州逸曜
            If glngModel = PM_住院医嘱清单 Or glngModel = PM_住院编辑 Then
                Select Case lngIndex
                Case HZYY_药品说明书
                     Call HZYY_DrugInstructions
                Case HZYY_药嘱审查
                    Call AdviceCheckWarn_HZYY(1, gobjPati.lng病人ID, gobjPati.str挂号单, gobjPati.lng主页ID)
                End Select
            ElseIf glngModel = PM_门诊编辑 Or glngModel = PM_门诊医嘱清单 Then
                Select Case lngIndex
                Case HZYY_药品说明书
                     Call HZYY_DrugInstructions
                Case HZYY_药嘱审查
                    Call AdviceCheckWarn_HZYY(1, gobjPati.lng病人ID, gobjPati.str挂号单, gobjPati.lng主页ID)
                End Select
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlPassAdviceDel(ByRef objMap As clsPassMap, ByVal lng医嘱ID As Long, ByVal dat当前时间 As Date)
'功能:医嘱删除
    Dim strXML As String
    
    If Not PassMap(objMap) Then Exit Sub
    
    If gbytPass = DT Then
        If gstrVersion = "3.0" Then
            If glngModel = PM_门诊医嘱清单 Then
                strXML = MakeMediDelXML(lng医嘱ID & "", dat当前时间)
                Call dtywzxUI(512, 0, strXML)
            End If
        End If
    End If
End Sub

Private Sub Class_Initialize()
    glngObject = glngObject + 1
End Sub

Private Sub Class_Terminate()
    '在导航台退出时激活,此时注销PASS
    Dim lngTmp As Long
    Dim strXML As String
    Dim strRet As Long
    
    glngObject = glngObject - 1
    If glngObject > 0 Then Exit Sub
    On Error Resume Next
    If PassModel = 1 Then Exit Sub
    If gbytPass = MK Then
        If gstrVersion = "3.0" Then
            Call PassQuit
        ElseIf gstrVersion = "4.0" Then
            Call MDC_Quit
        End If
        Set gcolPASSExe = Nothing
        Set gcolPASSState = Nothing
    ElseIf gbytPass = DT Then
        If gstrVersion = "3.0" Then
            lngTmp = dtywzxUI(3, 0, "") '清除灯的信息
            lngTmp = dtywzxUI(1, 0, "") '退出大通程序并关闭提醒灯
        ElseIf gstrVersion = "4.0" Then
            strXML = DTBS_MakeDetailXML(DTBS_退出)
            lngTmp = CRMS_UI(DTBS_退出, gstrBaseXml, strXML, "")
            gstrBaseXml = ""
        End If
    ElseIf gbytPass = TYT Then
        Set gobjPass = Nothing
    ElseIf gbytPass = YWS Then
        strXML = YWS_MakeDetailXML(YWS_退出)
        strRet = gobjPass.YWS_UI(YWS_退出, gstrBaseXml, strXML, "")
        gstrBaseXml = ""
    End If
    Set gobjPlugIn = Nothing
End Sub

Public Function zlPassAdviceSave(ByRef objMap As clsPassMap, Optional ByRef blnNoSave As Boolean, Optional ByVal bytFunc As Byte, _
    Optional ByVal strDelIDs As String) As Boolean
'功能:门诊、住院保存医嘱时调用审查接口
'参数:blnNoSave -是否保存 -True 需要将审查结果更新到数据库，False-无需将审查结果更新到数据库
'     Pass自动用药审查
'     bytFunc=1-停止医嘱,2-回退医嘱,3-作废,4-删除医嘱
'     strDelIDs -bytFunc=4 传入删除的医嘱ID（杭州逸曜）
    Dim blnIsHaveOut As Boolean   '审查药品是否存在离院带药
    Dim lngTmp As Long, lngReturn As Long
    Dim blnMsg As Boolean
    
    If Not PassMap(objMap) Then Exit Function
    
    If glngModel = PM_门诊编辑 Then
        If gbytPass = MK Then
            If gstrVersion = "3.0" Then
                blnMsg = OutAdviceCheckWarn_MK(MK_门诊保存审查, , blnNoSave) = 3
            Else
                blnMsg = OutAdviceCheckWarn_MK4(1, 1, , blnNoSave) = 1
            End If
            If blnMsg Then
                If gbytBlackLamp = 1 Then
                    If MsgBox("审查发现黑灯用药，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    Call MsgBox("审查发现黑灯用药，请作废或暂停黑灯用药后才能保存。", vbExclamation, gstrSysName)
                    Exit Function
                End If
            End If
        ElseIf gbytPass = DT Then
            If gstrVersion = "3.0" Then
                If OutAdviceCheckWarn_DT = False Then
                    Exit Function
                End If
            Else
                If AdviceCheckWarn_DTBS(1, False, , objMap) = False Then
                    Exit Function
                End If
            End If
        ElseIf gbytPass = YWS Then
            If OutAdviceCheckWarn_YWS(blnNoSave) = False Then
                Exit Function
            End If
        ElseIf gbytPass = TYT Then

            If OutAdviceCheckWarn_TYT(1, , blnNoSave) = 1 Then
                If gbytBlackLamp = 1 Then
                    If MsgBox("审查发现红灯用药，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    Call MsgBox("审查发现红灯用药，请作废或暂停红灯用药后才能保存。", vbExclamation, gstrSysName)
                    Exit Function
                End If
            End If
        ElseIf gbytPass = HZYY Then
            If bytFunc = 0 Then
                If AdviceCheckWarn_HZYY(0, gobjPati.lng病人ID, gobjPati.str挂号单, gobjPati.lng主页ID) = False Then
                    Exit Function
                End If
            End If
        End If
    ElseIf glngModel = PM_门诊医嘱清单 Then
        '门诊医嘱作废\删除 自动调用
        If gbytPass = MK And gstrVersion = "4.0" And InStr(",3,4,", "," & bytFunc & ",") > 0 Then
            Call OutAdviceCheckWarn_MK4(0, 1) '显示和采集
        ElseIf gbytPass = HZYY And bytFunc = 4 Then
            Call AdviceCheckWarn_HZYY(2, gobjPati.lng病人ID, gobjPati.str挂号单, gobjPati.lng主页ID, , strDelIDs)
        End If
    
    ElseIf glngModel = PM_住院编辑 Then
        If gbytPass = MK Then
            If gstrVersion = "3.0" Then
                blnMsg = InAdviceCheckWarn_MK(MK_住院保存审查, , blnIsHaveOut, blnNoSave) = 3
            Else
                blnMsg = InAdviceCheckWarn_MK4(1, 1, blnIsHaveOut, blnNoSave) = 1
            End If
            If blnMsg Then
                If gbytBlackLamp = 1 Then
                    If MsgBox("审查发现黑灯用药，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    If blnIsHaveOut And gbytOutBlackLamp = 1 Then
                        If MsgBox("存在院外执行的药品审查发现黑灯用药，是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            Exit Function
                        End If
                    Else
                        Call MsgBox("审查发现黑灯用药，请作废或暂停黑灯用药后才能保存。", vbExclamation, gstrSysName)
                        Exit Function
                    End If
                End If
            End If
        ElseIf gbytPass = DT Then
            If gstrVersion = "3.0" Then
                If InAdviceCheckWarn_DT = False Then
                    Exit Function
                End If
            Else
                If AdviceCheckWarn_DTBS(1, False, , objMap) = False Then
                    Exit Function
                End If
            End If
        ElseIf gbytPass = YWS Then
            If InAdviceCheckWarn_YWS(blnNoSave) = False Then
                Exit Function
            End If
        ElseIf gbytPass = TYT Then
            '审查通过保存病人病生状况
            If InAdviceCheckWarn_TYT(1, , blnIsHaveOut, blnNoSave) = 1 Then
                If gbytBlackLamp = 1 Then
                    If MsgBox("审查发现红灯用药，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    If blnIsHaveOut And gbytOutBlackLamp = 1 Then
                        If MsgBox("存在院外执行的药品审查发现红灯用药，是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            Exit Function
                        End If
                    Else
                        Call MsgBox("审查发现红灯用药，请作废或暂停红灯用药后才能保存。", vbExclamation, gstrSysName)
                        Exit Function
                    End If
                End If
            End If
        ElseIf gbytPass = HZYY Then
            If bytFunc = 0 Then
                If AdviceCheckWarn_HZYY(0, gobjPati.lng病人ID, gobjPati.str挂号单, gobjPati.lng主页ID) = False Then
                    Exit Function
                End If
            End If
        End If
    ElseIf glngModel = PM_住院医嘱清单 Then
        If gbytPass = MK And gstrVersion = "4.0" And InStr(",1,3,4,", "," & bytFunc & ",") > 0 Then
            '医嘱作废|停止|删除后自动调用
            Call InAdviceCheckWarn_MK4(0, 1) '显示和采集
        ElseIf gbytPass = MK And gstrVersion = "3.0" And InStr(",1,2,", "," & bytFunc & ",") > 0 Then
            Call InAdviceCheckWarn_MK(MK_手工调用审查)
        ElseIf gbytPass = HZYY And bytFunc = 4 Then
            Call AdviceCheckWarn_HZYY(3, gobjPati.lng病人ID, gobjPati.str挂号单, gobjPati.lng主页ID, , strDelIDs)
        End If
    End If
    zlPassAdviceSave = True
End Function

Public Sub zlPassAdviceInput(ByRef objMap As clsPassMap, ByVal lng诊疗项目ID As Long, ByVal lng收费细目ID As Long, Optional ByVal str药品名称 As String)
'功能:医嘱录入时合理用药检测
'参数:
    Dim strMsg As String
    Dim udtDetails As YWS_DETAILS
    Dim udtDTBSDetails As DTBS_DETAILS
    Dim strDetailsXml As String
    Dim strRetXML As String
    Dim strRet As String
    Dim strSQL As String
    Dim rs规格 As ADODB.Recordset
    Dim lngRet As Long
    
    If Not PassMap(objMap) Then Exit Sub
    If lng诊疗项目ID = 0 And lng收费细目ID = 0 Then Exit Sub
    
    If lng收费细目ID = 0 Then
        '按品种下达时,默认选择一个药品规格的药品id
        strSQL = "select 药品ID from 药品规格 where 药名id=[1] and rownum <2"
        Set rs规格 = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lng诊疗项目ID)
        lng收费细目ID = Val(rs规格!药品ID & "")
    End If
    
    If gbytPass = DT Then
        
        If gstrVersion = "3.0" Then
            If glngModel = PM_门诊编辑 Or glngModel = PM_住院编辑 Then
                strMsg = MakeMediXML(str药品名称, "" & lng收费细目ID)
                Call WriteLog(strMsg)
                strMsg = dtywzxUI(12, 0, strMsg) '要点提示
                Call WriteLog(strMsg)
            End If
        ElseIf gstrVersion = "4.0" Then
            If glngModel = PM_住院编辑 Or glngModel = PM_住院医嘱清单 Then
                udtDTBSDetails.str门诊住院标识 = "ip"
            Else
                udtDTBSDetails.str门诊住院标识 = "op"
            End If
            udtDTBSDetails.str药品代码 = lng收费细目ID
            udtDTBSDetails.str药品名称 = DTBS_StrToXML(str药品名称)
            strDetailsXml = DTBS_MakeMedicXML(udtDTBSDetails)
            Call WriteLog("功能号：" & DTBS_要点提示)
            Call WriteLog("BaseXml：" & gstrBaseXml)
            Call WriteLog("DetailsXml：" & strDetailsXml)
            lngRet = CRMS_UI(DTBS_要点提示, gstrBaseXml, strDetailsXml, strRetXML)
            Call WriteLog("strRetXML:" & strRetXML)
        End If
    ElseIf gbytPass = YWS Then
        If glngModel = PM_住院编辑 Or glngModel = PM_住院医嘱清单 Then
            udtDetails.str门诊住院标识 = "ip"
        Else
            udtDetails.str门诊住院标识 = "op"
        End If
        udtDetails.str药品代码 = lng收费细目ID
        udtDetails.str药品名称 = YWS_StrToXML(str药品名称)
        strDetailsXml = YWS_MakeMedicXML(udtDetails)
        strRet = gobjPass.YWS_UI(YWS_要点提示, gstrBaseXml, strDetailsXml, strRetXML)
    ElseIf gbytPass = MK And gstrVersion = "4.0" Then
        Call MDC_DoSetDrug(lng收费细目ID, str药品名称)
        Call MDC_DoRefDrug(51) '简要信息浮动窗口
        gbytOpen = FLOATWIN_DRUG
        glngDrugID = lng收费细目ID
    End If
End Sub

Public Sub zlPassCmdAlleyManage(ByRef objMap As clsPassMap, Optional ByVal blnDo As Boolean)
'功能：对病人过敏史/病生状态进行管理
'参数:
'    blnDo -True    PM_PIVA管理使用
    
    If Not PassMap(objMap) Then Exit Sub

    With gobjAdvice
        If gbytPass = MK Then
            If gstrVersion = "3.0" Then
                If glngModel = PM_门诊编辑 Then Call OutAdviceCheckWarn_MK(MK_病生状态过敏史)  '门诊编辑界面
                If glngModel = PM_住院编辑 Then Call InAdviceCheckWarn_MK(MK_病生状态过敏史)   '住院编辑界面
                '对病人过敏史/病生状态进行查看
                If glngModel = PM_护士校对 Then Call OperateAdviceCheckWarn_MK(MK_病生状态过敏史查看, .Row)  '护士校对界面
            Else
                Call frmPhysicalSel.ShowMe(objMap)
            End If
        ElseIf gbytPass = DT Or gbytPass = TYT Or gbytPass = YWS Then
            Call frmPhysicalSel.ShowMe(objMap)
        End If
    End With

End Sub

Public Sub zlPassAdviceColHidden(ByRef objMap As clsPassMap)
'功能:设置医嘱列是否隐藏
    If Not PassMap(objMap) Then Exit Sub
    gobjAdvice.ColHidden(gobjCOL.intCOL警示) = Not (InStr(",1,3,4,", "," & gbytPass & ",") > 0)    '大通隐藏警示列
End Sub

Public Sub zlPassCmdAlleyEnable(ByRef objMap As clsPassMap)
'功能:病生状态可用检测
'参数: udtPass -PASS接口对象映射
'       udtPP -病人信息
    '映射当前环境
    
    If Not PassMap(objMap) Then Exit Sub
    
    If gobjCmdAlley Is Nothing Then Exit Sub
    gobjCmdAlley.Visible = True
    gobjAdvice.ColHidden(gobjCOL.intCOL警示) = Not (InStr(",1,3,4,", "," & gbytPass & ",") > 0)    '大通隐藏警示列
    If glngModel = PM_门诊编辑 Or glngModel = PM_住院编辑 Then
        If gbytPass = MK Then
            If gstrVersion = "3.0" Then
                gobjCmdAlley.Enabled = PassGetState("AlleyEnable") = 1
            Else
                objMap.objCmdBar.Enabled = True
                gobjCmdAlley.Caption = "病生理情况(&B)"
            End If
        ElseIf gbytPass = DT Or gbytPass = YWS Then
            objMap.objCmdBar.Enabled = True
            gobjCmdAlley.Caption = "病生理情况(&B)"
        ElseIf gbytPass = TYT Then
            gobjCmdAlley.Enabled = True
            gobjCmdAlley.Caption = "症状/病生状态"
        ElseIf gbytPass = HZYY Or gbytPass = ZL Then
            gobjCmdAlley.Visible = False
        End If
    ElseIf glngModel = PM_护士校对 Then  '护士校对
        If gbytPass = MK Then
            If gstrVersion = "3.0" Then
                gobjCmdAlley.Enabled = PassGetState("AlleyEnable") = 1
            Else
                gobjCmdAlley.Visible = False
            End If
        End If
    End If
End Sub

Public Sub zlPassAdviceMainPoint(ByRef objMap As clsPassMap, Optional ByVal bytFunc As Byte = 0)
'功能:返回合理用药提示
'参数:bytFunc=调用场合0-双击事件,1-单击事件)
    Dim strMsg As String
    Dim strSQL As String
    Dim str药品ID As String
    Dim str药品名称 As String
    Dim str医嘱ID As String
    Dim strRetXML As String
    Dim strDetailsXml As String
    Dim strRet As String
    
    Dim udtDetails As YWS_DETAILS
    Dim udtDTBSDetails As DTBS_DETAILS
    Dim blnTag As Boolean
    Dim blnDo As Boolean
    
    Dim rs规格 As ADODB.Recordset
    
    Dim lngSelCol As Long
    Dim lngMouseRow As Long
    Dim lngMouseCol As Long
    Dim lngRet As Long
    
    If Not PassMap(objMap) Then Exit Sub
    
    On Error GoTo errH
    With gobjAdvice
        blnTag = False
        lngMouseRow = .MouseRow
        lngMouseCol = .MouseCol
        If lngMouseRow = -1 Then
            If gbytPass = MK And gstrVersion = "4.0" Then
                blnTag = True
            Else
                Exit Sub
            End If
        ElseIf InStr(",5,6,7,", .TextMatrix(.Row, objMap.VSCOL.intCOL诊疗类别)) > 1 Then
            str药品ID = .TextMatrix(.Row, objMap.VSCOL.intCOL收费细目ID) & ""
            str药品名称 = .TextMatrix(.Row, objMap.VSCOL.intCOL药品名称) & ""
            If bytFunc = 1 Then
                If gbytPass = MK And gstrVersion = "4.0" Then
                    If glngModel = PM_住院编辑 Or glngModel = PM_住院医嘱清单 Or glngModel = PM_门诊编辑 Or glngModel = PM_门诊医嘱清单 Or glngModel = PM_护士校对 Then
                        
                        lngSelCol = objMap.VSCOL.intCOL医嘱内容
                        If glngModel = PM_住院医嘱清单 Or glngModel = PM_门诊医嘱清单 Then
                            If .ColHidden(objMap.VSCOL.intCOL医嘱内容) = True Then lngSelCol = objMap.VSCOL.intCOL医嘱内容 + 1    '简介模式-显示内容列
                        End If
                        
                        If lngMouseCol = lngSelCol Then
                            If str药品ID = "" And .TextMatrix(.Row, objMap.VSCOL.intCOL诊疗项目ID) <> "" Then
                                strSQL = "select 药品ID from 药品规格 where 药名id=[1] and rownum <2"
                                Set rs规格 = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, .TextMatrix(.Row, objMap.VSCOL.intCOL诊疗项目ID))
                                str药品ID = rs规格!药品ID & ""
                            End If
                            Call MDC_DoSetDrug(str药品ID, str药品名称)
                            Call MDC_DoRefDrug(51) '简要信息提示
                            gbytOpen = FLOATWIN_DRUG
                        ElseIf lngMouseCol = objMap.VSCOL.intCOL警示 Then
                            If Not .Cell(flexcpPicture, lngMouseRow, lngMouseCol) Is Nothing Then
                                '警示列
                                If glngModel <> PM_护士校对 Then
                                    If glngModel = PM_住院编辑 Or glngModel = PM_门诊编辑 Then
                                        str医嘱ID = CStr(.RowData(.Row))
                                    Else
                                        str医嘱ID = .TextMatrix(.Row, objMap.VSCOL.intCOLID)
                                    End If
                                    Call MDC_ShowWarningHint(str医嘱ID)
                                    gbytOpen = FLOATWIN_WARN
                                End If
                            End If
                        Else
                            blnTag = True '点击空白也要关闭信息提示
                        End If
                    End If
                End If
            ElseIf bytFunc = 0 Then
                If gbytPass = DT Then
                    If gstrVersion = "3.0" Then
                        If glngModel = PM_门诊编辑 Or glngModel = PM_住院编辑 Or glngModel = PM_住院医嘱清单 Then
                            strMsg = MakeMediXML(Get项目名称(Val(.TextMatrix(.Row, objMap.VSCOL.intCOL诊疗项目ID))), .TextMatrix(.Row, objMap.VSCOL.intCOL收费细目ID), True)
                            Call WriteLog(strMsg)
                            strMsg = dtywzxUI(4108, 0, strMsg)    '要点提示
                            Call WriteLog(strMsg)
                        End If
                    ElseIf gstrVersion = "4.0" Then
                        If glngModel = PM_门诊编辑 Or glngModel = PM_住院编辑 Or glngModel = PM_住院医嘱清单 Or glngModel = PM_门诊医嘱清单 Then
                            If glngModel = PM_住院编辑 Or glngModel = PM_住院医嘱清单 Then
                                udtDTBSDetails.str门诊住院标识 = "ip"
                            Else
                                udtDTBSDetails.str门诊住院标识 = "op"
                            End If
                            udtDTBSDetails.str药品代码 = .TextMatrix(.Row, objMap.VSCOL.intCOL收费细目ID)
                            udtDTBSDetails.str药品名称 = DTBS_StrToXML(.TextMatrix(.Row, objMap.VSCOL.intCOL药品名称))
                            strDetailsXml = DTBS_MakeMedicXML(udtDTBSDetails)
                            WriteLog "功能号:" & DTBS_要点提示
                            WriteLog "DetailsXml：" & strDetailsXml
                            lngRet = CRMS_UI(DTBS_要点提示, gstrBaseXml, strDetailsXml, strRetXML)
                            WriteLog "RetValue:" & lngRet & vbCrLf & "RetXML：" & strRetXML
                        End If
                    End If
                ElseIf gbytPass = TYT Then
                    If glngModel = PM_门诊编辑 Then
                        If .Col = objMap.VSCOL.intCOL医嘱内容 Then
                            '药品提示
                            Call OutAdviceCheckWarn_TYT(TYT_药品提示, .Row)
                        End If
                        
                        If .Col = objMap.VSCOL.intCOL警示 And Not .Cell(flexcpPicture, .Row, .Col) Is Nothing Then  '太元通 双击击指示灯，调用窗口
                            '先审查后调用审查详情
                            Call OutAdviceCheckWarn_TYT(TYT_药嘱审查)
                            Call OutAdviceCheckWarn_TYT(TYT_审查详情)
                        End If
                    ElseIf glngModel = PM_住院编辑 Or glngModel = PM_住院医嘱清单 Then
                        '太元通 双击击指示灯，调用窗口
                        If .Col = objMap.VSCOL.intCOL警示 And Not .Cell(flexcpPicture, .Row, .Col) Is Nothing Then
                            Call InAdviceCheckWarn_TYT(TYT_药嘱审查)    '1-审查
                            Call InAdviceCheckWarn_TYT(TYT_审查详情)    '5-显示审查详情
                        End If
                        If .Col = objMap.VSCOL.intCOL医嘱内容 And glngModel = PM_住院医嘱清单 Then
                            Call InAdviceCheckWarn_TYT(TYT_药品提示, .Row)
                        End If
                    End If
                ElseIf gbytPass = YWS Then
                    If glngModel = PM_门诊编辑 Or glngModel = PM_住院编辑 Or glngModel = PM_住院医嘱清单 Or glngModel = PM_门诊医嘱清单 Then
                        If glngModel = PM_住院编辑 Or glngModel = PM_住院医嘱清单 Then
                            udtDetails.str门诊住院标识 = "ip"
                        Else
                            udtDetails.str门诊住院标识 = "op"
                        End If
                        udtDetails.str药品代码 = .TextMatrix(.Row, objMap.VSCOL.intCOL收费细目ID)
                        udtDetails.str药品名称 = .TextMatrix(.Row, objMap.VSCOL.intCOL药品名称)
                        strDetailsXml = YWS_MakeMedicXML(udtDetails)
                        strRet = gobjPass.YWS_UI(YWS_要点提示, gstrBaseXml, strDetailsXml, strRetXML)
                    End If
                End If
            End If
        Else
            '非药品医嘱时
            blnTag = True
        End If
        
        If blnTag Then
            If gbytPass = MK And gstrVersion = "4.0" And gbytOpen <> FLOATWIN_CLOSE Then
                If gbytOpen = FLOATWIN_DRUG Then
                    Call MDC_CloseDrugHint '点击空白也要关闭信息提示
                ElseIf gbytOpen = FLOATWIN_WARN Then
                    Call MDC_CloseWarningHint '点击空白也要关闭悬浮审查信息
                End If
                gbytOpen = FLOATWIN_CLOSE
            End If
        End If

    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlPassSetWarnLight(ByRef objMap As clsPassMap, ByVal lngRow As Long, ByVal int审查结果 As Integer)
'功能:界面加载医嘱时候设置警示灯
'参数:
'     lngRow:当前行
'     int审查结果 :审查结果值
    Dim arrLight(0 To 4) As String
    
    If Not PassMap(objMap) Then Exit Sub
    
    With gobjAdvice
        If gbytPass = MK Then
            If gstrVersion = "3.0" Then
                If InStr(",0,1,2,3,4,", int审查结果 & "") > 1 Then
                    .Cell(flexcpData, lngRow, gobjCOL.intCOL警示) = CStr(int审查结果)
                    Set .Cell(flexcpPicture, lngRow, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(int审查结果 + 1).Picture
                End If
            ElseIf gstrVersion = "4.0" Then
                If InStr(",0,1,2,3,4,", int审查结果 & "") > 1 Then
                    arrLight(0) = "蓝_4": arrLight(1) = "黑_4": arrLight(2) = "红_4": arrLight(3) = "橙_4": arrLight(4) = "黄_4"
                    .Cell(flexcpData, lngRow, gobjCOL.intCOL警示) = CStr(int审查结果)
                    Set .Cell(flexcpPicture, lngRow, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(int审查结果)).Picture
                End If
            End If
        ElseIf gbytPass = TYT Then       '太元通警示灯显示
            If InStr(",1,2,3,", int审查结果 & "") > 1 Then
                arrLight(1) = "红": arrLight(2) = "黄": arrLight(3) = "蓝"
                .Cell(flexcpData, lngRow, gobjCOL.intCOL警示) = CStr(int审查结果)
                Set .Cell(flexcpPicture, lngRow, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(int审查结果)).Picture
            End If
        ElseIf gbytPass = YWS Then
            If InStr(",0,1,2,3,", int审查结果 & "") > 1 Then
                arrLight(0) = "蓝": arrLight(1) = "黄": arrLight(2) = "橙": arrLight(3) = "红"
                .Cell(flexcpData, lngRow, gobjCOL.intCOL警示) = CStr(int审查结果)
                Set .Cell(flexcpPicture, lngRow, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(int审查结果)).Picture
            End If
        End If
    End With
End Sub

Public Sub zlPassShowWarn(ByRef objMap As clsPassMap)
'美康4.0
'参数号
    Dim strID As String
    
    If Not PassMap(objMap) Then Exit Sub
    
    With gobjAdvice
        Select Case glngModel
        Case PM_门诊编辑 To PM_住院首页
            If InStr(",5,6,7,", .TextMatrix(.Row, objMap.VSCOL.intCOL诊疗类别)) > 1 Then '大于1是为了避免诊疗类别为空串的情况
                If gbytPass = MK Then
                    If gstrVersion = "4.0" Then
                        If glngModel = PM_住院编辑 Or glngModel = PM_住院医嘱清单 Or glngModel = PM_门诊编辑 Or glngModel = PM_门诊医嘱清单 Then
                            If .Col = objMap.VSCOL.intCOL警示 And Not .Cell(flexcpPicture, .Row, .Col) Is Nothing Then  '单击警示灯
                                If glngModel = PM_住院编辑 Or glngModel = PM_门诊编辑 Then
                                    strID = .RowData(.Row)
                                Else
                                    strID = .TextMatrix(.Row, objMap.VSCOL.intCOLID)
                                End If
                                Call MDC_ShowWarningHint(strID)
                            End If
                        End If
                    End If
                End If
            End If
        End Select
    End With
End Sub

Public Sub zlPassShowWarn_YF(ByVal str医嘱ID As String)
'功能：处方审查系统调用,用于审查后,获取药品详情
'参数:  str医嘱ID-医嘱ID
'       blnShow -显示或关闭
    If gbytPass = MK Then
        If gstrVersion = "3.0" Then
            '获取PASS系统功能状态信息接口，判断PASS服务是否正常
            If PassGetState("PassEnable") = 0 Then
                MsgBox "当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
                Screen.MousePointer = 0: Exit Sub
            End If
            Call PassSetWarnDrug(str医嘱ID)         '传入药品查询信息接口
            Call PassDoCommand(MK_警示提示窗口)        '403警示色简要提示信息浮动窗口执行接口
            gbytOpen = FLOATWIN_WARN
        ElseIf gstrVersion = "4.0" Then
            Call MDC_ShowWarningHint(str医嘱ID)
            gbytOpen = FLOATWIN_WARN
        End If
    ElseIf gbytPass = TYT Then
        gobjPass.getDrugAlertDetail
    End If
End Sub

Public Sub zlPASSPopupCommandBars(ByRef objMap As clsPassMap, CommandBar As CommandBar, ByVal lngConMenu As Long)
'功能:frmInDockAdvice合理用药审查弹出菜单
'参数:gint场合 0-医生站，1-护士站，2-医技站
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Long, k As Long, lngIndex As Long
    Dim lngTmp As Long
    Dim blnDo As Boolean
    
    If Not PassMap(objMap) Then Exit Sub
    
    With CommandBar.Controls
        If Not CommandBar Is Nothing Then
            If gbytPass = DT Then
                CommandBar.Visible = False: Exit Sub
            Else
                CommandBar.Visible = True
            End If
        End If
        If Not CommandBar.Parent Is Nothing Then
            If Not (CommandBar.Parent.ID >= lngConMenu * 10# + 1 And CommandBar.Parent.ID <= lngConMenu * 10# + 99 Or CommandBar.Parent.ID = lngConMenu) Then
                Exit Sub
            End If
        End If
        
        blnDo = True
        If .Count > 0 Then '住院编辑，门诊编辑界面与住院医生站 弹出菜单 层级少一层
            For i = 1 To .Count
                If .Item(i).Parameter = "药嘱审查" Then blnDo = False: Exit For
            Next
        End If
        
        If gbytPass = MK Then
            If blnDo Then
                If gstrVersion = "3.0" Then '美康3.0
                    strTmp = "药物临床信息参考(&C),药品说明书(&D),中国药典(&N),病人用药教育(&S),检验值(&T),专项信息"
                    strTmp = strTmp & ",药物相互作用(&D),药食相互作用(&F),国内注射剂配伍(&M),国外注射剂配伍(&T),禁忌症(&C)"
                    strTmp = strTmp & ",副作用(&S),老年人用药(&G),儿童用药(&P),妊娠期用药(&E),哺乳期用药(&L)"
                    strTmp = strTmp & ",医药信息中心(&I),药品配对信息(&M),给药途径配对信息(&R),医院药品信息(&F)"
                    strTmp = strTmp & ",系统设置(&U),用药研究(&M),警告(&W),审查(&V)"
            
                    arrTmp = Split(strTmp, ",")
                    lngIndex = 0
                    For i = 0 To UBound(arrTmp)
                        If i >= 6 And i <= 15 Then '二级菜单
                            Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngConMenu * 10# + i, CStr(arrTmp(i)))
                            objControl.Parameter = "药嘱审查"
                        ElseIf i = 5 Then
                            lngIndex = lngIndex + 1
                            Set objPopup = .Add(xtpControlButtonPopup, lngConMenu * 10# + i, CStr(arrTmp(i)), lngIndex)
                            objPopup.Parameter = "药嘱审查"
                            objPopup.BeginGroup = True
                        ElseIf i < 5 Or i > 15 Then '一级菜单
                            lngIndex = lngIndex + 1
                            Set objControl = .Add(xtpControlButton, lngConMenu * 10# + i, CStr(arrTmp(i)), lngIndex)
                            objControl.Parameter = "药嘱审查"
                        End If
                        If i = 8 Or i = 10 Or i = 12 Or i = 16 Or i = 17 Or i = 20 Or i = 21 Then
                            objControl.BeginGroup = True
                        End If
                    Next
                    '将菜单索引与功能号映射到集合
                    Set gcolPASSExe = New Collection
                    With gcolPASSExe
                        .Add MK_药物临床信息参考, "_" & MK_IX_药物临床信息参考
                        .Add MK_药品说明书, "_" & MK_IX_药品说明书
                        .Add MK_中国药典, "_" & MK_IX_中国药典
                        .Add MK_病人用药教育, "_" & MK_IX_病人用药教育
                        .Add MK_检验值, "_" & MK_IX_检验值
                        'MK_IX_专项信息  二级标题名
                        .Add MK_药物_药物相互作用, "_" & MK_IX_药物相互作用
                        .Add MK_药物_食物相互使用, "_" & MK_IX_药食相互作用
                        .Add MK_国内注射剂配伍, "_" & MK_IX_国内注射剂配伍
                        .Add MK_国外注射剂配伍, "_" & MK_IX_国外注射剂配伍
                        .Add MK_禁忌症, "_" & MK_IX_禁忌症
                        .Add MK_副作用, "_" & MK_IX_副作用
                        .Add MK_老年人用药, "_" & MK_IX_老年人用药
                        .Add MK_儿童用药, "_" & MK_IX_儿童用药
                        .Add MK_妊娠期用药, "_" & MK_IX_妊娠期用药
                        .Add MK_哺乳期用药, "_" & MK_IX_哺乳期用药
                        .Add MK_医药信息中心, "_" & MK_IX_医药信息中心
                        .Add MK_药品配对信息, "_" & MK_IX_药品配对信息
                        .Add MK_给药途径配对信息, "_" & MK_IX_给药途径配对信息
                        .Add MK_医院药品信息, "_" & MK_IX_医院药品信息
                        .Add MK_系统设置, "_" & MK_IX_系统设置
                        .Add MK_用药研究, "_" & MK_IX_用药研究
                        .Add MK_单药警告, "_" & MK_IX_警告
                        .Add MK_手工调用审查, "_" & MK_IX_审查
                    End With
                    '美康菜单可用状态审查接口传人值与菜单索引的集合映射
                    Set gcolPASSState = New Collection
                    With gcolPASSState
                        .Add "CPRRes", "_" & MK_IX_药物临床信息参考
                        .Add "Directions", "_" & MK_IX_药品说明书
                        .Add "Chp", "_" & MK_IX_中国药典
                        .Add "CPERes", "_" & MK_IX_病人用药教育
                        .Add "CheckRes", "_" & MK_IX_检验值
                        'MK_IX_专项信息  二级标题名
                        .Add "DDIM", "_" & MK_IX_药物相互作用
                        .Add "DFIM", "_" & MK_IX_药食相互作用
                        .Add "MatchRes", "_" & MK_IX_国内注射剂配伍
                        .Add "TriessRes", "_" & MK_IX_国外注射剂配伍
                        .Add "DDCM", "_" & MK_IX_禁忌症
                        .Add "SIDE", "_" & MK_IX_副作用
                        .Add "GERI", "_" & MK_IX_老年人用药
                        .Add "PEDI", "_" & MK_IX_儿童用药
                        .Add "PREG", "_" & MK_IX_妊娠期用药
                        .Add "LACT", "_" & MK_IX_哺乳期用药
                        .Add "MEDInfo", "_" & MK_IX_医药信息中心
                        .Add "MATCH-DRUG", "_" & MK_IX_药品配对信息
                        .Add "MATCH-ROUTE", "_" & MK_IX_给药途径配对信息
                        .Add "HisDrugInfo", "_" & MK_IX_医院药品信息
                        .Add "SYS-SET", "_" & MK_IX_系统设置
                        .Add "DISQUISITION", "_" & MK_IX_用药研究
                    End With
                ElseIf gstrVersion = "4.0" Then '美康4.0
                    '提供了悬浮菜单
                    strTmp = "药嘱审查(&V)"
            
                    arrTmp = Split(strTmp, ",")
                    lngIndex = 0
                    For i = 0 To UBound(arrTmp)
                        If i < 5 Then  '一级菜单
                            lngIndex = lngIndex + 1
                            Set objControl = .Add(xtpControlButton, lngConMenu * 10# + i, CStr(arrTmp(i)), lngIndex)
                            objControl.Parameter = "药嘱审查"
                        End If
                    Next
                    '将菜单索引与功能号映射到集合
                    Set gcolPASSExe = New Collection
                    With gcolPASSExe
                        .Add MK4_审查, "_" & MK4_IX_审查
                    End With
                    
                End If
            End If
     
            If gstrVersion = "3.0" Then
                With gobjAdvice
                    If glngModel = PM_住院医嘱清单 Or glngModel = PM_住院编辑 Then
                        lngTmp = InAdviceCheckWarn_MK(MK_检测PASS菜单状态, .Row)
                    ElseIf glngModel = PM_门诊编辑 Or glngModel = PM_门诊医嘱清单 Then
                        lngTmp = OutAdviceCheckWarn_MK(MK_检测PASS菜单状态, .Row)
                    ElseIf glngModel = PM_护士校对 Then
                        lngTmp = OperateAdviceCheckWarn_MK(MK_检测PASS菜单状态, .Row)
                    End If
                End With
                For i = 1 To .Count
                    Call PassSetMenuState(.Item(i), gint场合, lngConMenu, lngTmp = 1)
                Next
            ElseIf gstrVersion = "4.0" Then
            
            End If
        ElseIf gbytPass = TYT Then    '太元通药嘱审查

            If blnDo Then
                strTmp = "用药规范(&0),药嘱审查(&1),药品提示(&2),医药知识库(&3),系统配置(&4)"
                arrTmp = Split(strTmp, ",")
                lngIndex = 0
                For i = 0 To UBound(arrTmp)
                    lngIndex = lngIndex + 1
                    Set objControl = .Add(xtpControlButton, lngConMenu * 10# + i, CStr(arrTmp(i)), lngIndex)
                    objControl.ID = lngConMenu * 10# + i
                    objControl.Parameter = "药嘱审查"
                Next
            End If
            
            For i = 1 To .Count
                Call PassSetMenuState(.Item(i), gint场合, lngConMenu, True)
            Next
        ElseIf gbytPass = HZYY Then
            If blnDo Then
                strTmp = "药品说明书(&0),药嘱审查(&1)"
                arrTmp = Split(strTmp, ",")
                lngIndex = 0
                For i = 0 To UBound(arrTmp)
                    lngIndex = lngIndex + 1
                    Set objControl = .Add(xtpControlButton, lngConMenu * 10# + i, CStr(arrTmp(i)), lngIndex)
                    objControl.ID = lngConMenu * 10# + i
                    objControl.Parameter = "药嘱审查"
                Next
            End If
            
            For i = 1 To .Count
                Call PassSetMenuState(.Item(i), gint场合, lngConMenu, True)
            Next
        End If
     End With
End Sub

Public Sub zlPassClearLight(ByRef objMap As clsPassMap)
'功能:清除警示灯信息
    Dim strRetXML As String
    Dim strRet As String
    
    If Not PassMap(objMap) Then Exit Sub
    
    If gbytPass = DT Then
        If InStr(",0,1,2,4,", glngModel) > 0 Then
            If gstrVersion = "3.0" Then
                Call dtywzxUI(3, 0, "")  '清除灯的信息
            ElseIf gstrVersion = "4.0" Then
                Call CRMS_UI(DTBS_初始UI, gstrBaseXml, "", strRetXML)
            End If
        End If
        
    ElseIf gbytPass = YWS Then
        strRet = gobjPass.YWS_UI(YWS_初始客户端, gstrBaseXml, "", strRetXML)
    End If
End Sub

Public Function zlPassCheck(ByRef objMap As clsPassMap) As Boolean
'功能:判断能不能调用合理用药监测接口
'参数: glngModel:调用模块号
'      gint场合:调用场合
'返回:True-能;False-不能
    Dim blnDo As Boolean
    
    If Not PassMap(objMap) Then Exit Function
    
    If gbytPass = 0 Then Exit Function
    '合理用药监测部件被多个模块公用，所以每次都要进行权限判断
    If Not PassCheckPrivs(glngModel) Then zlPassCheck = False: Exit Function

    blnDo = gbytPass > 0
    
    If blnDo Then
    '当操作员同时具备医生站和医技站的权限时，医生站启用Pass,医技站不启用Pass时,当先进入医生站，然后点开医技站，此时医技站也能启用PASS
        If gbytPass = MK Then
            If glngModel = PM_门诊编辑 Then
                If gint场合 = US_Intech Then blnDo = False
            End If
        ElseIf gbytPass = DT Then
            '即太元通和大通医技站、护士校对不使用合理用药。
            blnDo = (glngModel = PM_门诊编辑 Or glngModel = PM_住院编辑 Or (glngModel = PM_住院医嘱清单 And gint场合 = US_InDoctor) Or (glngModel = PM_门诊医嘱清单 And gint场合 = US_InDoctor))
        ElseIf gbytPass = TYT Then
            '这里分开写，是为了区别开，便于独立维护。
            blnDo = (glngModel = PM_门诊编辑 Or glngModel = PM_住院编辑 Or glngModel = PM_住院首页 Or ((glngModel = PM_住院医嘱清单 Or glngModel = PM_门诊医嘱清单) And gint场合 <> US_Intech))
        End If
    End If
    
    zlPassCheck = blnDo
End Function

Public Sub zlPassCommandBarUpdate(ByRef objMap As clsPassMap, ByRef control As CommandBarControl, ByVal blnAdvice As Boolean)
'功能:设置菜单状态
     
    'Call PassMap(udtPass, udtPP)
    '此处不需要再映射，住院医生站commandBar的Update事件为高频事件，从住院医生站进入医生编辑界面时(新开)，当医生编辑界面还未加载完时需要对过敏按钮状态进行设置，
    '而此时的update事件也在并发执行，导致住院编辑界面映射的过敏按钮对象 被此处的映射内容覆盖掉，报错 with block
    '因为没有映射，所以有些内容只能通过传人对象来取值
    If gbytPass = MK Or gbytPass = YWS Or gbytPass = DT Then
        If objMap.lngModel = PM_住院医嘱清单 Or objMap.lngModel = PM_门诊医嘱清单 Then
            control.Enabled = blnAdvice
            If blnAdvice Then
                control.Enabled = InStr(",5,6,7,", "," & objMap.vsAdvice.TextMatrix(objMap.vsAdvice.Row, objMap.VSCOL.intCOL诊疗类别) & ",") > 0
            End If
        End If
    ElseIf gbytPass = TYT Then
        '太元通不作处理 原因:当医嘱没有选中药品医嘱的时候，合理用药部分菜单项是可用状态 如系统设置
        
    End If
End Sub

Public Sub zlPassCommandBarAdd(ByRef objMap As clsPassMap, ByVal control As Object, ByVal lngID As Long, ByVal lngIconId As Long, Optional ByRef lngIdx As Long = -1)
'功能:添加菜单项目
'返回:lngIdx -工具栏需要返回索引值,使工具栏正确显示
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCmdBar As CommandBar
    
    If Not PassMap(objMap) Then Exit Sub
    With control
        If gbytPass = MK And gstrVersion = "3.0" Then
            If glngModel = PM_住院医嘱清单 Or glngModel = PM_门诊医嘱清单 Then
                If lngIdx = -1 Then  '菜单
                    Set objPopup = control.Add(xtpControlSplitButtonPopup, lngID, "药嘱审查(&K)")
                    objPopup.ID = lngID
                    objPopup.IconId = lngIconId
                    objPopup.BeginGroup = True
                Else          '工具栏
                    Set objPopup = control.Add(xtpControlSplitButtonPopup, lngID, "药嘱审查", lngIdx)
                    objPopup.ID = lngID
                    objPopup.IconId = lngIconId
                    lngIdx = objPopup.Index
                    objPopup.BeginGroup = True
                End If
            End If
        ElseIf gbytPass = MK And gstrVersion = "4.0" Then
            If glngModel = PM_住院医嘱清单 Or glngModel = PM_门诊医嘱清单 Then
                If gint场合 = US_InDoctor Then    '美康4.0接口 在护士站不设置 手动审查功能
                    If lngIdx = -1 Then
                        Set objControl = control.Add(xtpControlButton, lngID, "药嘱审查(&K)")
                        objControl.ID = lngID
                        objControl.IconId = lngIconId
                        objControl.BeginGroup = True
                    Else
                        Set objControl = control.Add(xtpControlButton, lngID, "药嘱审查", lngIdx)
                        objControl.ID = lngID
                        objControl.IconId = lngIconId
                        lngIdx = objControl.Index
                        objControl.BeginGroup = True
                    End If
                End If
            End If
        ElseIf gbytPass = DT Then
            If glngModel = PM_住院医嘱清单 Or glngModel = PM_门诊医嘱清单 Then
                If gint场合 = US_InDoctor Then    '大通接口 在护士站不需要设置 保存审查功能，故不添加
                    If lngIdx = -1 Then
                        Set objControl = control.Add(xtpControlButton, lngID, "药嘱审查(&K)")
                        objControl.ID = lngID
                        objControl.IconId = lngIconId
                        objControl.BeginGroup = True
                    Else
                        Set objControl = control.Add(xtpControlButton, lngID, "药嘱审查", lngIdx)
                        objControl.ID = lngID
                        objControl.IconId = lngIconId
                        lngIdx = objControl.Index
                        objControl.BeginGroup = True
                    End If
                End If
            End If
        ElseIf gbytPass = TYT Or gbytPass = HZYY Then
            If glngModel = PM_住院医嘱清单 Or glngModel = PM_门诊医嘱清单 Then
                If gint场合 = US_InDoctor Then
                    If lngIdx = -1 Then  '菜单
                        Set objPopup = control.Add(xtpControlSplitButtonPopup, lngID, "药嘱审查(&K)")
                        objPopup.ID = lngID
                        objPopup.IconId = lngIconId
                        objPopup.BeginGroup = True
                    Else          '工具栏
                        Set objPopup = control.Add(xtpControlSplitButtonPopup, lngID, "药嘱审查", lngIdx)
                        objPopup.ID = lngID
                        objPopup.IconId = lngIconId
                        lngIdx = objPopup.Index
                        objPopup.BeginGroup = True
                    End If
                End If
            End If
        ElseIf gbytPass = YWS Then
            If glngModel = PM_住院医嘱清单 Or glngModel = PM_门诊医嘱清单 Then
                If gint场合 = US_InDoctor Then
                    If lngIdx = -1 Then
                        Set objControl = control.Add(xtpControlButton, lngID, "药嘱审查(&K)")
                        objControl.ID = lngID
                        objControl.IconId = lngIconId
                        objControl.BeginGroup = True
                    Else
                        Set objControl = control.Add(xtpControlButton, lngID, "药嘱审查", lngIdx)
                        objControl.ID = lngID
                        objControl.IconId = lngIconId
                        lngIdx = objControl.Index
                        objControl.BeginGroup = True
                    End If
                End If
            End If
        End If
    End With
     
End Sub

Private Sub PassSetMenuState(ByVal control As CommandBarControl, ByVal int场合 As Integer, ByVal lngConNenu As Long, ByVal blnVisible As Boolean)
'功能：设置Pass功能 的可见性及可用性
'说明:
    Dim lngTmp As Long
    Dim strDrugCode As String
    Dim lngQuery As Long
    Dim blnExit As Boolean
    
    control.Visible = blnVisible
    control.Enabled = blnVisible
    '避免没有医嘱时点鼠标右键,弹出无效菜单（空白小方块）
    If control.Visible = False And (glngModel = PM_门诊编辑 Or glngModel = PM_住院编辑) Then
        control.Visible = True
        control.Enabled = False
        blnExit = True
    End If
    
    If control.Visible And Not blnExit Then
        lngTmp = control.ID - lngConNenu * 10#
        If gbytPass = MK Then
            If gstrVersion = "3.0" Then
                If glngModel = PM_住院医嘱清单 And int场合 = US_Intech Then
                    If lngTmp = MK_IX_药品说明书 Then
                        control.Visible = True
                    Else
                        control.Visible = False
                    End If
                ElseIf glngModel = PM_住院医嘱清单 And int场合 <> US_Intech Then
                    control.Visible = True
                ElseIf glngModel = PM_护士校对 Then
                    If lngTmp > MK_IX_系统设置 Then control.Visible = False
                ElseIf glngModel = PM_PIVA管理 Or glngModel = PM_部门发药 Then
                    If lngTmp >= MK_IX_系统设置 Then control.Visible = False
                ElseIf glngModel = PM_处方发药 Then
                    If lngTmp = MK_IX_系统设置 Or lngTmp = MK_IX_用药研究 Then control.Visible = False
                End If
                
                If lngTmp = MK_IX_系统设置 Then '系统设置
                    control.Visible = control.Visible And (gbytSysSet = 1) 'gbytSysSet =0 隐藏，=1显示
                End If
            Else
                If glngModel = PM_住院医嘱清单 And (int场合 = US_InNurse Or int场合 = US_Intech) Then
                    If lngTmp = MK4_IX_药品说明书 Then
                        control.Visible = True
                    Else
                        control.Visible = False
                    End If
                ElseIf glngModel = PM_住院医嘱清单 And int场合 = US_InDoctor Then
                    control.Visible = True
                ElseIf glngModel = PM_护士校对 Then
                    If lngTmp > MK4_IX_细菌耐药率 Then control.Visible = False
                End If
            End If
            
            If control.Visible Then
                If gstrVersion = "3.0" Then
                    Select Case lngTmp
                    Case MK_IX_药物临床信息参考 To MK_IX_检验值, MK_IX_药物相互作用 To MK_IX_用药研究
                        control.Enabled = PassGetState(CStr(gcolPASSState("_" & lngTmp))) = 1
                    Case MK_IX_警告
                        '警告:有警示值(不为空),且大于0-蓝灯
                        If glngModel = PM_处方发药 Then
                            control.Enabled = Val(gobjAdvice.TextMatrix(gobjAdvice.Row, gobjCOL.intCOL警示)) > 0
                        Else
                            control.Enabled = Val(gobjAdvice.Cell(flexcpData, gobjAdvice.Row, gobjCOL.intCOL警示)) > 0
                        End If
                    End Select
                ElseIf gstrVersion = "4.0" Then '美康4.0
                    With gobjAdvice
                        Select Case lngTmp
                        Case MK4_IX_药品说明书 To MK4_IX_药品简要信息, MK4_IX_药物相互作用 To MK4_IX_细菌耐药率
                            strDrugCode = .TextMatrix(.Row, gobjCOL.intCOL收费细目ID)
                            If Val(strDrugCode) = 0 Then '收费细目ID为0则取 YMID & 诊疗项目ID
                                strDrugCode = "YMID" & .TextMatrix(.Row, gobjCOL.intCOL诊疗项目ID)
                            End If
                            control.Enabled = MDC_GetDrugRefEnabled(strDrugCode, gcolPASSExe("_" & lngTmp)) > 0
                        Case MK4_IX_系统设置  '系统设置
                            control.Visible = (gbytSysSet = 1) 'gbytSysSet =0 隐藏，=1显示
                        End Select
                    End With
                End If
            End If
        
        ElseIf gbytPass = TYT Then  '太元通菜单状态设置
            If int场合 = US_InDoctor Or int场合 = US_InNurse Then
                Select Case lngTmp
                
                Case TYT_用药规范   '用药规范
                    control.Enabled = True
                Case TYT_药嘱审查    '药嘱审查
                    control.Enabled = True
                Case TYT_药品提示   '药品提示
                    control.Enabled = InStr(",5,6,7,", "," & gobjAdvice.TextMatrix(gobjAdvice.Row, gobjCOL.intCOL诊疗类别) & ",") > 0
                Case TYT_医药知识库    '医药知识库
                    control.Enabled = True
                Case TYT_系统配置     '系统配置
                    control.Enabled = True
                End Select
            End If
        ElseIf gbytPass = HZYY Then  '杭州逸曜
            If int场合 = US_InDoctor Or int场合 = US_InNurse Then
                If int场合 = US_InNurse Then
                    If lngTmp = HZYY_药品说明书 Then
                        control.Visible = True
                    Else
                        control.Visible = False
                    End If
                End If
                If control.Visible Then
                    Select Case lngTmp
                    
                    Case HZYY_药品说明书
                        control.Enabled = True
                    Case HZYY_药嘱审查    '药嘱审查
                        control.Enabled = True
                    End Select
                End If
            End If
        End If
    End If
    
    '门诊/住院医嘱清单管理界面菜单Update事件需要
    control.Category = IIf(control.Visible, ";可见;", ";不可见;") & IIf(control.Enabled, ";可用;", ";不可用;")
End Sub

Public Function zlPassInputAllergy() As String
'功能:录入过敏源
    If gbytPass = TYT Then
        zlPassInputAllergy = gobjPass.inputAllergy
    End If
End Function

Private Function PassMap(ByRef objMap As clsPassMap) As Boolean
'功能:传人接口映射对象
'参数:
    '当调用场合发生变动,需要重新映射

    glngModel = objMap.lngModel
    If glngModel = 1 Then
        gbytUseType = objMap.bytUseType
    End If

    Set gobjCOL = objMap.VSCOL

    Set gfrmMain = objMap.frmMain
    Set gobjAdvice = objMap.vsAdvice
    Set gobjCmdAlley = objMap.objCmdBar
    If glngModel = PM_门诊编辑 Then
        Set gobjDiags = objMap.Diags
    End If
    gint场合 = objMap.int场合
    '病人信息
    Set gobjPati = objMap.PassPati
    
    If Not gblnInitOK Then
        If Not PassInitialize Then
            gbytPass = UNPASS
            PassType = UNPASS
            gblnInitOK = False
            Exit Function
        End If
        gblnInitOK = True
    End If
    PassMap = True
End Function

Public Sub zlPassSetDrug(ByRef objMap As clsPassMap)
'功能：行列变换,传药品信息给用户
    Dim rs规格 As ADODB.Recordset
    Dim strSQL As String
    Dim str药品ID  As String
    
    If Not PassMap(objMap) Then Exit Sub
    With gobjAdvice
        Select Case glngModel
        Case PM_住院编辑, PM_住院医嘱清单, PM_门诊编辑, PM_门诊医嘱清单, PM_护士校对
            If gbytPass = MK Then
                If gstrVersion = "4.0" Then
                    str药品ID = .TextMatrix(.Row, objMap.VSCOL.intCOL收费细目ID)
                    If glngDrugID = Val(str药品ID) Then Exit Sub '相同药品不必重复调用
                    
                    If gbytOpen = FLOATWIN_DRUG Then
                        Call MDC_CloseDrugHint '关闭简要信息提示'点击空白也要关闭信息提示
                    ElseIf gbytOpen = FLOATWIN_WARN Then
                        Call MDC_CloseWarningHint
                    End If
                    If Not InStr(",5,6,7,", .TextMatrix(.Row, objMap.VSCOL.intCOL诊疗类别)) > 1 Then
                        Call MDC_DoSetDrug("", "")    '清空记录
                        glngDrugID = 0
                        Exit Sub
                    End If
                    
                    If str药品ID = "" And .TextMatrix(.Row, objMap.VSCOL.intCOL诊疗项目ID) <> "" Then
                        strSQL = "select 药品ID from 药品规格 where 药名id=[1] and rownum <2"
                        Set rs规格 = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, .TextMatrix(.Row, objMap.VSCOL.intCOL诊疗项目ID))
                        str药品ID = rs规格!药品ID & ""
                    End If
                    Call MDC_DoSetDrug(str药品ID, .TextMatrix(.Row, objMap.VSCOL.intCOL药品名称))
                    glngDrugID = CLng(Val(str药品ID))
                End If
            End If
        End Select
    End With
End Sub

Public Sub zlPassUpLoad(ByRef objMap As clsPassMap)
    If Not PassMap(objMap) Then Exit Sub
    If gbytPass = DT Then
        If gstrVersion = "4.0" Then
            Call AdviceCheckWarn_DTBS(1, True, , objMap)
        End If
    ElseIf gbytPass = HZYY Then
        Call AdviceCheckWarn_HZYY(4, gobjPati.lng病人ID, gobjPati.str挂号单, gobjPati.lng主页ID)
    End If
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'---------药房调用接口
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub PassSetMenuState_YF(ByVal control As CommandBarControl, ByVal lngConNenu As Long, ByVal blnVisible As Boolean, _
             Optional ByVal str警示 As String)
'功能：设置Pass功能 的可见性及可用性
'说明:
    Dim lngTmp As Long
    Dim strDrugCode As String
    Dim lngQuery As Long

    
    control.Visible = True    '避免所有都为false
    control.Enabled = blnVisible
    
    If blnVisible Then
        lngTmp = control.ID - lngConNenu * 10#
        
        If gbytPass = MK Then
            If blnVisible Then
                If gstrVersion = "3.0" Then
                    Select Case lngTmp
                    Case MK_IX_药物临床信息参考 To MK_IX_检验值, MK_IX_药物相互作用 To MK_IX_用药研究
                       
                        If lngTmp = MK_IX_系统设置 Or lngTmp = MK_IX_用药研究 Then
                            '处方发药,部门发药,输液配置中心都不提供这两个选项
                            control.Visible = False
                        Else
                            control.Enabled = PassGetState(CStr(gcolPASSState("_" & lngTmp))) = 1
                        End If
                    Case MK_IX_警告
                        If glngModel = PM_处方发药 Then
                            '警告:有警示值(不为空),且大于0-蓝灯
                            control.Enabled = Val(str警示) > 0
                        ElseIf glngModel = PM_部门发药 Then
                            control.Enabled = Val(str警示) > 0
                        ElseIf glngModel = PM_PIVA管理 Then
                            control.Visible = False
                        End If
                        
                    Case MK_IX_审查
                        If glngModel = PM_处方发药 Then
                            control.Visible = True
                        ElseIf glngModel = PM_部门发药 Then
                            control.Visible = True
                        ElseIf glngModel = PM_PIVA管理 Then
                            control.Visible = True
                        End If
                    End Select
                ElseIf gstrVersion = "4.0" Then '美康4.0
                    
                
                End If
            End If
        ElseIf gbytPass = TYT Then  '太元通菜单状态设置
            Select Case lngTmp
            Case TYT_用药规范   '用药规范
                control.Enabled = True
            Case TYT_药嘱审查    '药嘱审查
                control.Enabled = True
            Case TYT_药品提示   '药品提示
                control.Enabled = True
            Case TYT_医药知识库    '医药知识库
                control.Enabled = True
            Case TYT_系统配置     '系统配置
                control.Enabled = True
            End Select
        ElseIf gbytPass = HZYY Then  '杭州逸曜菜单状态设置
            Select Case lngTmp
            Case HZYY_药品说明书    '药品说明书
                control.Enabled = True
            Case HZYY_药嘱审查    '药嘱审查
                control.Enabled = True
            End Select
        End If
    End If
    
    '门诊/住院医嘱清单管理界面菜单Update事件需要
    control.Category = IIf(control.Visible, ";可见;", ";不可见;") & IIf(control.Enabled, ";可用;", ";不可用;")
End Sub


Public Sub zlPassClearLight_YF()
'功能:清除警示灯信息
    Dim strRetXML As String
    Dim strRet As String
    
    If gbytPass = DT Then
        If gstrVersion = "3.0" Then
            Call dtywzxUI(3, 0, "")  '清除灯的信息
        ElseIf gstrVersion = "4.0" Then
            Call CRMS_UI(DTBS_初始UI, gstrBaseXml, "", "")
        End If
    ElseIf gbytPass = YWS Then
        strRet = gobjPass.YWS_UI(YWS_初始客户端, gstrBaseXml, "", strRetXML)
    End If
End Sub

Public Function zlPassColHidden_YF() As Boolean
'功能:用于药房设置警示列的显示或隐藏
'大通、药卫士隐藏警示列
    zlPassColHidden_YF = Not (InStr(",1,2,3,4,", "," & gbytPass & ",") > 0)
End Function

Public Sub zlPassCommandBarAdd_YF(ByVal lngModel As Long, ByVal control As Object, ByVal lngID As Long, _
                            ByVal lngIconId As Long, Optional ByRef lngIdx As Long = -1)
'功能:添加菜单项目
'参数:
'返回:lngIdx -工具栏需要返回索引值,使工具栏正确显示
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCmdBar As CommandBar
    
    With control
        If gbytPass = MK Then
            If gstrVersion = "3.0" Then
                If lngModel = PM_PIVA管理 Then
                    If lngIdx = -1 Then
                        Set objControl = control.Add(xtpControlButton, lngID, "过敏史/病生状态(&G)")
                        objControl.ID = lngID
                        objControl.IconId = lngIconId
                    Else
                        Set objControl = control.Add(xtpControlButton, lngID, "过敏史/病生状态", lngIdx)
                        objControl.ID = lngID
                        objControl.IconId = lngIconId
                        lngIdx = objControl.Index
                    End If
                ElseIf lngModel = PM_处方发药 Then
                   
                    Set objCmdBar = control.Add("条件", xtpBarTop)
                    objCmdBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
                    objCmdBar.ModifyStyle XTP_CBRS_GRIPPER, 0
                    objCmdBar.ContextMenuPresent = False
    
                    Set objControl = objCmdBar.Controls.Add(xtpControlButton, lngID, "过敏史/病生状态")
                    objControl.BeginGroup = True
                    objControl.ToolTipText = "提示：显示过敏史/病生状态"
                    objControl.Style = xtpButtonIconAndCaption
                    objControl.IconId = lngIconId
                
                ElseIf lngModel = PM_部门发药 Then
                    Set objControl = control.Add(xtpControlButton, lngID, "过敏史/病生状态")
                    objControl.BeginGroup = True
                    objControl.ToolTipText = "提示：显示过敏史/病生状态"
                    objControl.Style = xtpButtonIconAndCaption
                    objControl.IconId = lngIconId
                End If
            Else
                
            End If
        ElseIf gbytPass = DT Then
          
        ElseIf gbytPass = TYT Then
            
        ElseIf gbytPass = YWS Then
'            If lngIdx = -1 Then
'                Set objControl = control.Add(xtpControlButton, lngID, "药嘱审查(&K)")
'                objControl.ID = lngID
'                objControl.IconId = lngIconId
'                objControl.BeginGroup = True
'            Else
'                Set objControl = control.Add(xtpControlButton, lngID, "药嘱审查", lngIdx)
'                objControl.ID = lngID
'                objControl.IconId = lngIconId
'                lngIdx = objControl.Index
'                objControl.BeginGroup = True
'            End If
        End If
    End With
     
End Sub

Public Function zlPassInit_YF(cnMain As ADODB.Connection, ByVal lngSys As Long, Optional ByVal lngModel As Long) As Boolean
'功能:用于药房合理用药监测接口初始化
'参数:lngModel

    glngSys = lngSys
    glngModel = lngModel
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    Set gcnOracle = cnMain
    
        
    '初始化
    If Not InitSysPar Then Exit Function
    Call GetUserInfo
    
    '设置PASS类型
    PassType = gbytPass
    PassVersion = gstrVersion
    '初始化失败
    If Not PassInitialize Then
        gbytPass = UNPASS
        PassType = UNPASS
        Exit Function
    End If
    
    zlPassInit_YF = True
End Function

Public Sub zlPASSPopupCommandBars_YF(ByVal lngModel As Long, CommandBar As CommandBar, ByVal lngConMenu As Long, ByVal lng病人ID As Long, _
                                     ByVal lng主页ID As Long, ByVal str挂号单 As String, Optional ByVal str警示 As String, Optional ByVal lngAdviceID As Long)
'功能:frmInDockAdvice合理用药审查弹出菜单
'参数:gint场合 0-医生站，1-护士站，2-医技站
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Long, k As Long, lngIndex As Long
    Dim lngTmp As Long
    Dim blnDo As Boolean
    
    glngModel = lngModel
    
    With CommandBar.Controls
        If Not CommandBar Is Nothing Then CommandBar.Visible = True
        If Not CommandBar.Parent Is Nothing Then
            If Not (CommandBar.Parent.ID >= lngConMenu * 10# + 1 And CommandBar.Parent.ID <= lngConMenu * 10# + 99 Or CommandBar.Parent.ID = lngConMenu) Then
                Exit Sub
            End If
        End If
        
        blnDo = True
        If .Count > 0 Then '住院编辑，门诊编辑界面与住院医生站 弹出菜单 层级少一层
            For i = 1 To .Count
                If .Item(i).Parameter = "药嘱审查" Then blnDo = False: Exit For
            Next
        End If
        
        If gbytPass = MK Then
            If blnDo Then
                If gstrVersion = "3.0" Then '美康3.0
                    strTmp = "药物临床信息参考(&C),药品说明书(&D),中国药典(&N),病人用药教育(&S),检验值(&T),专项信息"
                    strTmp = strTmp & ",药物相互作用(&D),药食相互作用(&F),国内注射剂配伍(&M),国外注射剂配伍(&T),禁忌症(&C)"
                    strTmp = strTmp & ",副作用(&S),老年人用药(&G),儿童用药(&P),妊娠期用药(&E),哺乳期用药(&L)"
                    strTmp = strTmp & ",医药信息中心(&I),药品配对信息(&M),给药途径配对信息(&R),医院药品信息(&F)"
                    strTmp = strTmp & ",系统设置(&U),用药研究(&M),警告(&W),审查(&V)"
                    
                    arrTmp = Split(strTmp, ",")
                    lngIndex = 0
                    For i = 0 To UBound(arrTmp)
                        If i >= 6 And i <= 15 Then '二级菜单
                            Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngConMenu * 10# + i, CStr(arrTmp(i)))
                            objControl.Parameter = "药嘱审查"
                        ElseIf i = 5 Then
                            lngIndex = lngIndex + 1
                            Set objPopup = .Add(xtpControlButtonPopup, lngConMenu * 10# + i, CStr(arrTmp(i)), lngIndex)
                            objPopup.Parameter = "药嘱审查"
                            objPopup.BeginGroup = True
                        ElseIf i < 5 Or i > 15 Then '一级菜单
                            lngIndex = lngIndex + 1
                            Set objControl = .Add(xtpControlButton, lngConMenu * 10# + i, CStr(arrTmp(i)), lngIndex)
                            objControl.Parameter = "药嘱审查"
                        End If
                        If i = 8 Or i = 10 Or i = 12 Or i = 16 Or i = 17 Or i = 20 Or i = 21 Then
                            objControl.BeginGroup = True
                        End If
                    Next
                    '将菜单索引与功能号映射到集合
                    Set gcolPASSExe = New Collection
                    With gcolPASSExe
                        .Add MK_药物临床信息参考, "_" & MK_IX_药物临床信息参考
                        .Add MK_药品说明书, "_" & MK_IX_药品说明书
                        .Add MK_中国药典, "_" & MK_IX_中国药典
                        .Add MK_病人用药教育, "_" & MK_IX_病人用药教育
                        .Add MK_检验值, "_" & MK_IX_检验值
                        'MK_IX_专项信息  二级标题名
                        .Add MK_药物_药物相互作用, "_" & MK_IX_药物相互作用
                        .Add MK_药物_食物相互使用, "_" & MK_IX_药食相互作用
                        .Add MK_国内注射剂配伍, "_" & MK_IX_国内注射剂配伍
                        .Add MK_国外注射剂配伍, "_" & MK_IX_国外注射剂配伍
                        .Add MK_禁忌症, "_" & MK_IX_禁忌症
                        .Add MK_副作用, "_" & MK_IX_副作用
                        .Add MK_老年人用药, "_" & MK_IX_老年人用药
                        .Add MK_儿童用药, "_" & MK_IX_儿童用药
                        .Add MK_妊娠期用药, "_" & MK_IX_妊娠期用药
                        .Add MK_哺乳期用药, "_" & MK_IX_哺乳期用药
                        .Add MK_医药信息中心, "_" & MK_IX_医药信息中心
                        .Add MK_药品配对信息, "_" & MK_IX_药品配对信息
                        .Add MK_给药途径配对信息, "_" & MK_IX_给药途径配对信息
                        .Add MK_医院药品信息, "_" & MK_IX_医院药品信息
                        .Add MK_系统设置, "_" & MK_IX_系统设置
                        .Add MK_用药研究, "_" & MK_IX_用药研究
                        .Add MK_单药警告, "_" & MK_IX_警告
                        .Add MK_手工调用审查, "_" & MK_IX_审查
                    End With
                    '美康菜单可用状态审查接口传人值与菜单索引的集合映射
                    Set gcolPASSState = New Collection
                    With gcolPASSState
                        .Add "CPRRes", "_" & MK_IX_药物临床信息参考
                        .Add "Directions", "_" & MK_IX_药品说明书
                        .Add "Chp", "_" & MK_IX_中国药典
                        .Add "CPERes", "_" & MK_IX_病人用药教育
                        .Add "CheckRes", "_" & MK_IX_检验值
                        'MK_IX_专项信息  二级标题名
                        .Add "DDIM", "_" & MK_IX_药物相互作用
                        .Add "DFIM", "_" & MK_IX_药食相互作用
                        .Add "MatchRes", "_" & MK_IX_国内注射剂配伍
                        .Add "TriessRes", "_" & MK_IX_国外注射剂配伍
                        .Add "DDCM", "_" & MK_IX_禁忌症
                        .Add "SIDE", "_" & MK_IX_副作用
                        .Add "GERI", "_" & MK_IX_老年人用药
                        .Add "PEDI", "_" & MK_IX_儿童用药
                        .Add "PREG", "_" & MK_IX_妊娠期用药
                        .Add "LACT", "_" & MK_IX_哺乳期用药
                        .Add "MEDInfo", "_" & MK_IX_医药信息中心
                        .Add "MATCH-DRUG", "_" & MK_IX_药品配对信息
                        .Add "MATCH-ROUTE", "_" & MK_IX_给药途径配对信息
                        .Add "HisDrugInfo", "_" & MK_IX_医院药品信息
                        .Add "SYS-SET", "_" & MK_IX_系统设置
                        .Add "DISQUISITION", "_" & MK_IX_用药研究
                    End With
                ElseIf gstrVersion = "4.0" Then '美康4.0
                    '提供了悬浮菜单
                    strTmp = "药嘱审查(&V)"
            
                    arrTmp = Split(strTmp, ",")
                    lngIndex = 0
                    For i = 0 To UBound(arrTmp)
                        If i < 5 Then  '一级菜单
                            lngIndex = lngIndex + 1
                            Set objControl = .Add(xtpControlButton, lngConMenu * 10# + i, CStr(arrTmp(i)), lngIndex)
                            objControl.Parameter = "药嘱审查"
                        End If
                    Next
                    '将菜单索引与功能号映射到集合
                    Set gcolPASSExe = New Collection
                    With gcolPASSExe
                        .Add MK4_审查, "_" & MK4_IX_审查
                    End With
                End If
            End If
            
            If gstrVersion = "3.0" Then
                lngTmp = AdviceCheckWarn_MK_YF(lng病人ID, lng主页ID, str挂号单, MK_检测PASS菜单状态, lngAdviceID)
                For i = 1 To .Count
                    Call PassSetMenuState_YF(.Item(i), lngConMenu, lngTmp = 1, str警示)
                Next
            ElseIf gstrVersion = "4.0" Then
                
            End If
        ElseIf gbytPass = TYT Then    '太元通药嘱审查

            If blnDo Then
                strTmp = "用药规范(&0),药嘱审查(&1),药品提示(&2),医药知识库(&3),系统配置(&4)"
                arrTmp = Split(strTmp, ",")
                lngIndex = 0
                For i = 0 To UBound(arrTmp)
                    lngIndex = lngIndex + 1
                    Set objControl = .Add(xtpControlButton, lngConMenu * 10# + i, CStr(arrTmp(i)), lngIndex)
                    objControl.ID = lngConMenu * 10# + i
                    objControl.Parameter = "药嘱审查"
                Next
            End If
            
            For i = 1 To .Count
                Call PassSetMenuState_YF(.Item(i), lngConMenu, True)
            Next
        ElseIf gbytPass = YWS Then
            '提供了悬浮菜单
            If blnDo Then
                strTmp = "药嘱审查(&V)"
        
                arrTmp = Split(strTmp, ",")
                lngIndex = 0
                For i = 0 To UBound(arrTmp)
                    If i < 5 Then  '一级菜单
                        lngIndex = lngIndex + 1
                        Set objControl = .Add(xtpControlButton, lngConMenu * 10# + i, CStr(arrTmp(i)), lngIndex)
                        objControl.Parameter = "药嘱审查"
                    End If
                Next
            End If
        ElseIf gbytPass = DT Then
            If blnDo Then
                strTmp = "药嘱审查(&V)"
                arrTmp = Split(strTmp, ",")
                lngIndex = 0
                For i = 0 To UBound(arrTmp)
                    lngIndex = lngIndex + 1
                    Set objControl = .Add(xtpControlButton, lngConMenu * 10# + i, CStr(arrTmp(i)), lngIndex)
                    objControl.Parameter = "药嘱审查"
                Next
            End If
        ElseIf gbytPass = HZYY Then    '杭州逸曜药嘱审查
            If blnDo Then
                strTmp = "用药规范(&0),药嘱审查(&1)"
                arrTmp = Split(strTmp, ",")
                lngIndex = 0
                For i = 0 To UBound(arrTmp)
                    lngIndex = lngIndex + 1
                    Set objControl = .Add(xtpControlButton, lngConMenu * 10# + i, CStr(arrTmp(i)), lngIndex)
                    objControl.ID = lngConMenu * 10# + i
                    objControl.Parameter = "药嘱审查"
                Next
            End If
            
            For i = 1 To .Count
                Call PassSetMenuState_YF(.Item(i), lngConMenu, True)
            Next
        End If
        
     End With
End Sub

Public Function zlPassSetWarnLight_YF(ByVal int审查结果 As Integer) As Object
'功能:根据审查结果值返回警示灯图片
'参数:int审查结果 -审查结果值
'返回：图片对象 (根据审查结果找不到图片时返回Nothing)
    Dim arrLight(0 To 4) As String
    Dim objPic As Object
    
    If gbytPass = MK Then
        If gstrVersion = "3.0" Then
            If InStr(",0,1,2,3,4,", int审查结果 & "") > 1 Then
                Set objPic = frmIcons.imgPass.ListImages(int审查结果 + 1).Picture
            End If
        Else
            If InStr(",0,1,2,3,4,", int审查结果 & "") > 1 Then
                arrLight(0) = "蓝_4": arrLight(1) = "黑_4": arrLight(2) = "红_4": arrLight(3) = "橙_4": arrLight(4) = "黄_4"
                Set objPic = frmIcons.imgPass.ListImages(arrLight(int审查结果)).Picture
            End If
        End If
    ElseIf gbytPass = TYT Then       '太元通警示灯显示
        If InStr(",1,2,3,", int审查结果 & "") > 1 Then
            arrLight(1) = "红": arrLight(2) = "黄": arrLight(3) = "蓝"
            Set objPic = frmIcons.imgPass.ListImages(arrLight(int审查结果)).Picture
        End If
    ElseIf gbytPass = YWS Then
        If InStr(",0,1,2,3,", int审查结果 & "") > 1 Then
            arrLight(0) = "蓝": arrLight(1) = "黄": arrLight(2) = "橙": arrLight(3) = "红"
            Set objPic = frmIcons.imgPass.ListImages(arrLight(int审查结果)).Picture
        End If
    End If
    Set zlPassSetWarnLight_YF = objPic
End Function

Public Sub zlPassQueryCheckResult_YF(ByVal lngModel As Long, ByVal strID As String, ByVal strFlag As String)
'功能:获取审查详情
'参数: 接口类型：TYT:strID-住院号\门诊号,strFlag：1-门诊\2-住院

    Select Case lngModel
    
    Case PM_处方发药, PM_部门发药, PM_PIVA管理
        If gbytPass = TYT Then
            Call gobjPass.queryCheckResult(strID, strFlag)
        End If
    End Select
End Sub

Public Sub zlPassCommandBarExe_YF(ByVal lngModel As Long, ByVal lngIndex As Long, ByVal lng病人ID As Long, _
                                 ByVal lng主页ID As Long, ByVal str挂号单 As String, Optional ByVal lngCurrAdviceID As Long)
'功能:药房合理用药监测执行
'参数:
'lngCurrAdviceID -当前某行医嘱ID
    Dim lngTmp As Long
    
    On Error GoTo errH
    glngModel = lngModel
    If gbytPass = MK Then    '美康审查
        If gstrVersion = "3.0" Then
            Select Case lngIndex
            Case MK_IX_药物临床信息参考 To MK_IX_检验值, MK_IX_药物相互作用 To MK_IX_系统设置
                Call PassDoCommand(gcolPASSExe("_" & lngIndex))
            Case MK_IX_用药研究
                '用药研究 药房用药研究不可见,系统设置不可见
            Case MK_IX_警告
                '警告:有警示值(不为空),且大于0-蓝灯
                Call AdviceCheckWarn_MK_YF(lng病人ID, lng主页ID, str挂号单, MK_单药警告, lngCurrAdviceID)
                '审查
            Case MK_IX_审查
                Call AdviceCheckWarn_MK_YF(lng病人ID, lng主页ID, str挂号单, MK_手工调用审查, lngCurrAdviceID)
            End Select
        ElseIf gstrVersion = "4.0" Then
            Call AdviceCheckWarn_MK4_YF(lng病人ID, lng主页ID, str挂号单)
        End If
    ElseIf gbytPass = DT Then   '大通
        If gstrVersion = "3.0" Then
            Call AdviceCheckWarn_DT(lng病人ID, lng主页ID, str挂号单)
        Else
            Call AdviceCheckWarn_DTBS(2, False, , , lng病人ID, lng主页ID, str挂号单)
        End If
    ElseIf gbytPass = TYT Then   '太元通药嘱审查
        Call AdviceCheckWarn_TYT_YF(lng病人ID, lng主页ID, str挂号单, lngIndex, lngCurrAdviceID)
    ElseIf gbytPass = YWS Then '药卫士
        Call AdviceCheckWarn_YWS_YF(lng病人ID, lng主页ID, str挂号单)
    ElseIf gbytPass = HZYY Then
    
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlPassCmdAlleyManage_YF(ByVal lngModel As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal str挂号单 As String)
'功能：对病人过敏史/病生状态进行管理
'参数:
'    blnDo -True    PM_PIVA管理使用
    
    With gobjAdvice
        If gbytPass = MK Then
            If InStr(",1341,1342,1345,", lngModel) > 0 Then
                Call AdviceCheckWarn_MK_YF(lng病人ID, lng主页ID, str挂号单, MK_病生状态过敏史查看)
            End If
        End If
    End With

End Sub

Public Sub zlPassAdviceMainPoint_YF(ByVal strFlag As String, ByVal str药品ID As String, ByVal str药品名称 As String, Optional ByVal bytFunc As Byte, Optional ByVal str医嘱ID As String)
'功能:双击医嘱行，返回合理用药提示
'参数:strFlag-1:门诊,2-住院
'    str药品ID-药品ID
'    str药品名称-药品名称
'    bytFunc=0 美康接口简要信息提示,1-关闭简要信息,2-显示审查结果,3-关闭悬浮审查信息
    Dim strDetailsXml As String
    Dim strRetXML As String
    Dim udtDetails As YWS_DETAILS
    Dim udtDTBSDetails As DTBS_DETAILS
    Dim strRet As String
    Dim strMsg As String
    Dim lngRet As Long
            
    If gbytPass = YWS Then
        If strFlag = "2" Then
            udtDetails.str门诊住院标识 = "ip"
        Else
            udtDetails.str门诊住院标识 = "op"
        End If
        udtDetails.str药品代码 = str药品ID
        udtDetails.str药品名称 = YWS_StrToXML(str药品名称)
        strDetailsXml = YWS_MakeMedicXML(udtDetails)
        strRet = gobjPass.YWS_UI(YWS_要点提示, gstrBaseXml, strDetailsXml, strRetXML)

    ElseIf gbytPass = MK And gstrVersion = "4.0" Then
        If bytFunc = 0 Then
            Call MDC_DoSetDrug(str药品ID, str药品名称)
            Call MDC_DoRefDrug(51)  '显示简要信息
        ElseIf bytFunc = 1 Then
            Call MDC_CloseDrugHint '点击空白也要关闭信息提示
        ElseIf bytFunc = 2 Then
            Call MDC_ShowWarningHint(str医嘱ID) '显示审查结果
        ElseIf bytFunc = 3 Then
            Call MDC_CloseWarningHint '点击空白也要关闭悬浮审查信息
        End If
    ElseIf gbytPass = DT Then
        If gstrVersion = "3.0" Then
            strMsg = MakeMediXML(str药品名称, str药品ID, True)
            Call WriteLog(strMsg)
            strMsg = dtywzxUI(4108, 0, strMsg)    '要点提示
        ElseIf gstrVersion = "4.0" Then
            If strFlag = "2" Then
                udtDTBSDetails.str门诊住院标识 = "ip"
            Else
                udtDTBSDetails.str门诊住院标识 = "op"
            End If
            udtDTBSDetails.str药品代码 = str药品ID
            udtDTBSDetails.str药品名称 = DTBS_StrToXML(str药品名称)
            strDetailsXml = DTBS_MakeMedicXML(udtDTBSDetails)
            WriteLog "功能号:" & DTBS_要点提示 & vbCrLf & "DetailXML:" & strDetailsXml
            lngRet = CRMS_UI(YWS_要点提示, gstrBaseXml, strDetailsXml, strRetXML)
            WriteLog "retValue:" & lngRet & vbCrLf & "RetXML:" & strRetXML
        End If
    End If
    
End Sub

Public Sub zlPassSetDrug_YF(ByVal str药品ID As String, ByVal str药品名称 As String)
'功能:向悬浮窗体传人药品信息
    If gbytPass = MK Then
        If gstrVersion = "4.0" Then
            Call MDC_CloseDrugHint '关闭简要信息提示'点击空白也要关闭信息提示
            Call MDC_DoSetDrug("", "")    '清空记录
            If str药品ID <> "" Then
                Call MDC_DoSetDrug(str药品ID, str药品名称)
            End If
        End If
    End If
End Sub

Public Sub zlPassCloseDrugHint(ByRef objMap As clsPassMap)
'功能:用于医嘱编辑关闭悬浮窗体
    If Not PassMap(objMap) Then Exit Sub
    If gbytPass = MK Then
        If gstrVersion = "4.0" And gbytOpen = FLOATWIN_DRUG Then
            Call MDC_CloseDrugHint '关闭简要信息提示'点击空白也要关闭信息提示
            gbytOpen = FLOATWIN_CLOSE
        End If
    End If
End Sub

Public Sub zlPassRecipelCheck(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
            ByVal str挂号单 As String, ByVal str医嘱IDs As String, Optional str警示 As String = "-1")
'功能:处方审查系统调用
'PM_门诊处方审查 ,PM_住院药嘱审查
'a)  "处方审查系统"调用PASS接口，并传入医嘱ID（药嘱），得到合理用药结果值（integer/byte）；
'b)  "处方审查系统"调用PASS接口，并传入医嘱ID（药嘱），能够显示相关的用药信息窗体（PASS商提供的窗体）；

    If gbytPass = MK Then
        If gstrVersion = "3.0" Then
            Call AdviceCheckWarn_MK_YF(lng病人ID, lng主页ID, str挂号单, MK_手工调用审查, , str医嘱IDs, str警示)
        Else
            Call AdviceCheckWarn_MK4_YF(lng病人ID, lng主页ID, str挂号单, str医嘱IDs, str警示)
        End If
    ElseIf gbytPass = DT Then
        If gstrVersion = "3.0" Then
            Call AdviceCheckWarn_DT(lng病人ID, lng主页ID, str挂号单, str医嘱IDs)
        Else
            Call AdviceCheckWarn_DTBS(2, False, , , lng病人ID, lng主页ID, str挂号单, str医嘱IDs)
        End If
        str警示 = ""
    ElseIf gbytPass = TYT Then
        Call AdviceCheckWarn_TYT_YF(lng病人ID, lng主页ID, str挂号单, 1, , str警示, str医嘱IDs)
    ElseIf gbytPass = YWS Then
        Call AdviceCheckWarn_YWS_YF(lng病人ID, lng主页ID, str挂号单, str医嘱IDs, str警示)
        str警示 = ""
    End If

End Sub

Public Sub Setup(objFrm As Object, cnMain As ADODB.Connection, ByVal lngSys As Long)
'设置参数
    Dim frmPara As New frmSetPara
    On Error GoTo errH

    glngSys = lngSys
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    Set gcnOracle = cnMain
    PassModel = 1
    
    frmPara.Show 1, objFrm
    
    Exit Sub
errH:
    MsgBox "参数设置失败！错误出现在:第" & CStr(Erl()) & "行," & Err.Description, vbExclamation, "中联信息"
End Sub

Public Function SetEnabled(ByVal bytPass As Byte, ByVal strVersion As String) As Boolean
'功能:检查是否允许提供电子签名参数界面
    Dim blnDo As Boolean
    gbytPass = bytPass
    gstrVersion = strVersion
    PassModel = 1
    
    Select Case bytPass
    Case DT
        If gstrVersion = "4.0" Then
            blnDo = True
        End If
    Case MK
        If gstrVersion = "4.0" Then
            blnDo = False   '干预系统35.10上
        End If
    Case HZYY
        blnDo = True
    End Select
    
    SetEnabled = blnDo
End Function
