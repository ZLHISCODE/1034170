VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPeisPDFReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjPeisReport As Object
Private mblnPrintOver As Boolean
Private mblnPrintDept As Boolean
Private mblnPrintSum As Boolean
Private mblnPrintEmpty As Boolean
Private mblnPrintQuestion As Boolean
Private mobjReport As Object
Private mstrSQL As String
Private mstrTitle As String

'######################################################################################################################
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'功能：输出指定人员的PDF体检报告单
'参数：
'返回：
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long)
    
    Set gcnOracle = cnOracle
    
'    ParamInfo.系统名称 = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
'    ParamInfo.产品名称 = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
        
'    UserInfo.用户名 = strDBUser
'    SetDbUser UserInfo.用户名
    
'    Call InitCommon(gcnOracle)
    
    glngSys = lngSys
    mblnPrintOver = (Val(gobjDatabase.GetPara("打印封面", lngSys, lngModual, "1")) >= 1)
    mblnPrintDept = (Val(gobjDatabase.GetPara("打印项目", lngSys, lngModual, "1")) >= 1)
    mblnPrintSum = (Val(gobjDatabase.GetPara("打印总检", lngSys, lngModual, "1")) >= 1)
    mblnPrintEmpty = (Val(gobjDatabase.GetPara("打印空项", lngSys, lngModual, "1")) >= 1)
    mblnPrintQuestion = (Val(gobjDatabase.GetPara("打印问卷", lngSys, lngModual, "1")) >= 1)
                                                     
    Call SaveSetting("ZLSOFT", "公共模块\zl9PeisComLib\个人健康体检报告", "打印机", Printer.DeviceName)
    
    
    Set mobjReport = CreateObject("zl9Report.clsReport")
    
    Initialize = True
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'功能：输出指定人员的PDF体检报告单
'参数：
'返回：
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function ExportPDFReport(ByVal lngTaskKey As Long, ByVal lngPersonKey As Long, Optional ByVal strPath As String) As Boolean
    Dim rsStyle As New ADODB.Recordset
        
    Set rsStyle = Get_Peistasktypeitem(lngTaskKey)
    If rsStyle.RecordCount > 0 Then
        Do While Not rsStyle.EOF
            
            Select Case Val(rsStyle("报告种类").Value)
            Case 3                  '专用体检报告
                Call OutputSpecialReport(rsStyle, lngTaskKey, lngPersonKey, strPath)
            Case Else               '普通体检报告
                
                Call OutputNormalReport(Val(rsStyle("ID").Value), lngTaskKey, lngPersonKey, strPath)
            End Select
            
            rsStyle.MoveNext
        Loop
    End If
End Function

'######################################################################################################################
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'功能：输出指定人员的体检报告单
'参数：
'返回：
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Function OutputNormalReport(ByVal lngReportStyleKey As Long, ByVal lngTaskKey As Long, ByVal lngPersonKey As Long, Optional ByVal strFile As String)
    
    If mobjPeisReport Is Nothing Then Set mobjPeisReport = CreateObject("zl9PeisComLib.clsPeisReport")
    
    If mobjPeisReport.InitReport(gcnOracle, lngReportStyleKey) Then
        If mobjPeisReport.GetPersonReportData(lngTaskKey, lngPersonKey, Me, mblnPrintOver, 5) Then
        
            If strFile <> "" Then
                
                OutputNormalReport = mobjPeisReport.ExportReport(3, strFile)
            End If
        End If
    End If
            
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'功能：输出指定人员的体检报告单
'参数：
'返回：
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Function OutputSpecialReport(ByVal rsReportStyle As ADODB.Recordset, ByVal lngTaskKey As Long, ByVal lngPatientKey As Long, ByVal strFile As String) As Boolean

    Dim blnNoPrint As Boolean
'    Dim strExcelName As String
'    Dim strExcelFile As String
'    Dim mobjclsPeisReport As New clsPeisReport
    Dim strTmp As String
    Dim varTmp As Variant
    Dim intCount As Integer
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
'    Dim rsGroup As New ADODB.Recordset
    Dim str报告格式 As String
    Dim rsData As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    Dim strClassName As String
    Dim str组件属性 As String
    Dim lngCount As Long
    Dim strUrlStart As String
    Dim aryProperty As Variant
    Dim intLoop As Integer
    Dim intPos As Integer
    Dim strPDFFile As String
    Dim intFileNum As Integer
    Dim strFileName As String
    Dim strFileExt As String
    
    On Error GoTo errHand
        
    If mobjPeisReport Is Nothing Then Set mobjPeisReport = CreateObject("zl9PeisComLib.clsPeisReport")
    
    Set rsData = New ADODB.Recordset
    With rsData
        .Fields.Append "类型", adTinyInt
        .Fields.Append "ID", adBigInt
        .Fields.Append "格式", adVarChar, 50
        .Fields.Append "封面", adTinyInt
        .Fields.Append "总检", adTinyInt
        .Fields.Append "项目", adTinyInt
        .Open
    End With
    
    If rsReportStyle.RecordCount > 0 Then
'        rsReportStyle.MoveFirst
        
        Set rsCondition = CreateCondition
        Call SetCondition(rsCondition, "报告组件id", gobjComFun.NVL(rsReportStyle("封面组件id").Value, 0))
        Set rsTmp = GetPeisComponet(2, rsCondition)
        If rsTmp.BOF = False Then
            strClassName = UCase(gobjComFun.NVL(rsTmp("类名").Value))
            str组件属性 = gobjComFun.NVL(rsTmp("组件属性").Value)
            
            If InStr(str组件属性, "报表格式=") > 0 Then str报告格式 = Mid(str组件属性, InStr(str组件属性, "报表格式=") + 5)
            If InStr(str组件属性, "输出内容=") > 0 Then str报告格式 = Mid(str组件属性, InStr(str组件属性, "输出内容=") + 5)
                            
            Select Case strClassName
            '--------------------------------------------------------------------------------------------------
            Case "ZL21_BILL_2121_4_1"
                rsData.AddNew
                rsData("类型").Value = 2
                rsData("ID").Value = 0
                rsData("格式").Value = str报告格式
                rsData("封面").Value = 1
                rsData("总检").Value = 0
                rsData("项目").Value = 0
            '--------------------------------------------------------------------------------------------------
            Case "ZL21_BILL_2121_4_3"
                rsData.AddNew
                rsData("类型").Value = 2
                rsData("ID").Value = 0
                rsData("格式").Value = str报告格式
                rsData("封面").Value = 0
                rsData("总检").Value = 1
                rsData("项目").Value = 0
            '--------------------------------------------------------------------------------------------------
            Case "ZL21_BILL_2121_4_2"
                rsData.AddNew
                rsData("类型").Value = 2
                rsData("ID").Value = 0
                rsData("格式").Value = str报告格式
                rsData("封面").Value = 0
                rsData("总检").Value = 0
                rsData("项目").Value = 1
            Case Else
                rsData.AddNew
                rsData("类型").Value = 1
                rsData("ID").Value = gobjComFun.NVL(rsReportStyle("ID").Value, 0)
                rsData("格式").Value = str报告格式
                rsData("封面").Value = 1
                rsData("总检").Value = 0
                rsData("项目").Value = 0
            End Select
        End If
        
        Call SetCondition(rsCondition, "报告目录id", Val(rsReportStyle("报告目录id").Value))
        Set rsTmp = GetPeisReportContent(3, rsCondition)
        If rsTmp.BOF = False Then
            Do While Not rsTmp.EOF
                strClassName = UCase(gobjComFun.NVL(rsTmp("类名").Value))
                
                str组件属性 = Trim(gobjComFun.NVL(rsTmp("组件属性").Value))
                aryProperty = Split(str组件属性, "'")
                For intLoop = 0 To UBound(aryProperty)
                    strTmp = CStr(aryProperty(intLoop))
                    intPos = InStr(strTmp, "=")
                    If intPos > 0 Then
                        Select Case Trim(Left(strTmp, intPos - 1))
                        Case "报表格式"
                            str报告格式 = Trim(Mid(strTmp, intPos + 1))
                        Case "输出内容"
                            str报告格式 = Trim(Mid(strTmp, intPos + 1))
                        Case "下载路径前缀"
                            If str报告格式 = "项目" Then
                                strUrlStart = Trim(Mid(strTmp, intPos + 1))
                            End If
                        End Select
                    End If
                Next
                Select Case strClassName
                '--------------------------------------------------------------------------------------------------
                Case "ZL21_BILL_2121_4_3"
                
                    rsData.Filter = ""
                    rsData.Filter = "类型=2 And 格式='" & str报告格式 & "'"
                    If rsData.RecordCount = 0 Then
                        rsData.AddNew
                        rsData("类型").Value = 2
                        rsData("ID").Value = 0
                        rsData("格式").Value = str报告格式
                        rsData("封面").Value = 0
                        rsData("总检").Value = 1
                        rsData("项目").Value = 0
                    Else
                        rsData("总检").Value = 1
                    End If

                '--------------------------------------------------------------------------------------------------
                Case "ZL21_BILL_2121_4_2"
                    rsData.Filter = ""
                    rsData.Filter = "类型=2 And 格式='" & str报告格式 & "'"
                    If rsData.RecordCount = 0 Then
    
                        rsData.AddNew
                        rsData("类型").Value = 2
                        rsData("ID").Value = 0
                        rsData("格式").Value = str报告格式
                        rsData("封面").Value = 0
                        rsData("总检").Value = 0
                        rsData("项目").Value = 1
    
                    Else
                        rsData("项目").Value = 1
                    End If
                '--------------------------------------------------------------------------------------------------
                Case Else
                    rsData.Filter = ""
                    rsData.Filter = "类型=1 And 格式='" & strClassName & "'"

                    Select Case str报告格式
                    Case "问卷"
                        If rsData.RecordCount = 0 Then
                            rsData.AddNew
                            rsData("类型").Value = 1
                            rsData("ID").Value = gobjComFun.NVL(rsReportStyle("ID").Value, 0)
                            rsData("格式").Value = strClassName
                            rsData("封面").Value = 0
                            rsData("总检").Value = 0
                            rsData("项目").Value = 1
                        Else
                            rsData("项目").Value = 1
                        End If
                    Case "总检"
                        If rsData.RecordCount = 0 Then
                            rsData.AddNew
                            rsData("类型").Value = 1
                            rsData("ID").Value = gobjComFun.NVL(rsReportStyle("ID").Value, 0)
                            rsData("格式").Value = strClassName
                            rsData("封面").Value = 0
                            rsData("总检").Value = 1
                            rsData("项目").Value = 0
                        Else
                            rsData("总检").Value = 1
                        End If
                    Case "项目"
                        If rsData.RecordCount = 0 Then
                            rsData.AddNew
                            rsData("类型").Value = 1
                            rsData("ID").Value = gobjComFun.NVL(rsReportStyle("ID").Value, 0)
                            rsData("格式").Value = strClassName
                            rsData("封面").Value = 0
                            rsData("总检").Value = 0
                            rsData("项目").Value = 1
                        Else
                            rsData("项目").Value = 1
                        End If
                    End Select
                    
                End Select
                
                rsTmp.MoveNext
            Loop
        End If
        
        
        '--------------------------------------------------------------------------------------------------------------
        rsData.Filter = ""
        rsData.Sort = "封面 Desc,总检 Desc,项目 Desc"
        If rsData.RecordCount > 0 Then rsData.MoveFirst
        
        lngCount = 0
        '同一报告组内的报告格式不同
        Do While Not rsData.EOF
            lngCount = lngCount + 1
            str报告格式 = gobjComFun.NVL(rsData("格式").Value)
            
            If Val(gobjComFun.NVL(rsData("类型").Value)) = 1 Then
                '固定报告
                If mobjPeisReport.InitReport(gcnOracle, Val(rsData("ID").Value)) Then
                    
                    
                    Call GetFileNameExt(strFile, strFileName, strFileExt)
                    strPDFFile = IIf(intFileNum = 0, strFileName, strFileName & intFileNum) & "." & strFileExt
                    intFileNum = intFileNum + 1
                    
                    Call mobjPeisReport.ExportReportSpec(lngTaskKey, lngPatientKey, (Val(rsData("封面").Value) = 1 And mblnPrintOver), _
                                                            (Val(rsData("项目").Value) = 1 And mblnPrintDept), _
                                                            (Val(rsData("总检").Value) = 1 And mblnPrintSum), _
                                                            mblnPrintEmpty, 5, _
                                                            "", strPDFFile, True, _
                                                            (mblnPrintQuestion And lngCount = 1), _
                                                            strUrlStart)
                End If
            Else
                Call mobjPeisReport.InitReport(gcnOracle, Val(rsData("ID").Value))
                
                intCount = intCount + 1
                
                '自定义报表
                '------------------------------------------------------------------------------------------------------
                If Val(rsData("封面").Value) = 1 And mblnPrintOver Then
                    Set rs = Get_Rptfmtno("ZL21_BILL_2121_4_1", str报告格式)
                    If rs.BOF = False Then str报告格式 = rs("序号").Value
                    
                    If Val(str报告格式) > 0 Then
                    
                        Call GetFileNameExt(strFile, strFileName, strFileExt)
                        strPDFFile = IIf(intFileNum = 0, strFileName, strFileName & intFileNum) & "." & strFileExt
                        intFileNum = intFileNum + 1
                    
                        Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL21_BILL_2121_4_1", Me, "任务id=" & lngTaskKey, "病人id=" & lngPatientKey, "REPORTFORMAT=" & Val(str报告格式), "PDF=" & strPDFFile, "PrintEmpty=0", 4)
                    End If
                End If
                
                '------------------------------------------------------------------------------------------------------
                If Val(rsData("总检").Value) = 1 And mblnPrintSum Then
                    Set rs = Get_Rptfmtno("ZL21_BILL_2121_4_3", str报告格式)
                    If rs.BOF = False Then str报告格式 = rs("序号").Value
                    
                    If Val(str报告格式) > 0 Then
                        Call GetFileNameExt(strFile, strFileName, strFileExt)
                        strPDFFile = IIf(intFileNum = 0, strFileName, strFileName & intFileNum) & "." & strFileExt
                        intFileNum = intFileNum + 1
                    
                        Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL21_BILL_2121_4_3", Me, "任务id=" & lngTaskKey, "病人id=" & lngPatientKey, "REPORTFORMAT=" & Val(str报告格式), "PDF=" & strPDFFile, "PrintEmpty=0", 4)
                    End If
                End If
                
                '------------------------------------------------------------------------------------------------------
                If Val(rsData("项目").Value) = 1 And mblnPrintDept Then

                    Set rs = Get_Rptfmtno("ZL21_BILL_2121_4_2", str报告格式)
                    If rs.BOF = False Then str报告格式 = rs("序号").Value
                    
                    If Val(str报告格式) > 0 Then
                        Call GetFileNameExt(strFile, strFileName, strFileExt)
                        strPDFFile = IIf(intFileNum = 0, strFileName, strFileName & intFileNum) & "." & strFileExt
                        intFileNum = intFileNum + 1
                    
                        Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL21_BILL_2121_4_2", Me, "任务id=" & lngTaskKey, "病人id=" & lngPatientKey, "空项=" & IIf(mblnPrintEmpty, 1, 0), "REPORTFORMAT=" & Val(str报告格式), "PDF=" & strPDFFile, "PrintEmpty=0", 4)
                    End If
                    
                End If
                                
            End If
            
            rsData.MoveNext
        Loop
    End If
        
    OutputSpecialReport = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function Get_Peistasktypeitem(ByVal 任务ID_In As Long, Optional ByVal lngType As Long, Optional ByVal 场合_In As Long = 0) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    Dim n_体检类型id As Long
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand

    mstrSQL = "Select Nvl(体检类型id,0) As 体检类型id From 体检任务记录 Where ID=[1]"
    
'    If gclsBase.IsDataMoved(DataMovedMode.PeisTaskKey, 任务ID_In) = True Then
'        mstrSQL = gclsBase.ReplaceBakTable(mstrSQL, "体检任务记录")
'    End If
    
    Set rs = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, 任务ID_In)
    If rs.BOF = False Then n_体检类型id = rs("体检类型id").Value
    If lngType > 0 Then
        mstrSQL = _
        "Select b.ID,b.报告种类,a.报告目录id,a.报告组号,b.报告编码,b.报告名称,b.封面组件id From 体检类型报告 a,体检报告目录 b Where a.体检类型id=[1] And a.报告目录id=b.Id And b.报告种类 =[2]  Order By a.报告组号"
    Else
        If 场合_In = 0 Then
            mstrSQL = _
                "Select b.ID,b.报告种类,a.报告目录id,a.报告组号,b.报告编码,b.报告名称,b.封面组件id From 体检类型报告 a,体检报告目录 b Where a.体检类型id=[1] And a.报告目录id=b.Id And b.报告种类 In (1,3) Order By a.报告组号"
        ElseIf 场合_In = 1 Then
            mstrSQL = _
                "Select b.ID,b.报告种类,a.报告目录id,a.报告组号,b.报告编码,b.报告名称,b.封面组件id From 体检类型报告 a,体检报告目录 b Where a.体检类型id=[1] And a.报告目录id=b.Id And b.报告种类 =4 Order By a.报告组号"
        ElseIf 场合_In = 2 Then
            mstrSQL = _
                "Select b.ID,b.报告种类,a.报告目录id,a.报告组号,b.报告编码,b.报告名称,b.封面组件id From 体检类型报告 a,体检报告目录 b Where a.体检类型id=[1] And a.报告目录id=b.Id And b.报告种类 =1 Order By a.报告组号"
        
        End If
    End If
    Set Get_Peistasktypeitem = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, n_体检类型id, lngType)

    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetPeisComponet(ByVal bytMode As Byte, ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    On Error GoTo errHand
    '报告组件
    Select Case bytMode
    Case 1
        mstrSQL = _
            "Select '组件' As 图标,a.ID,Decode(a.组件种类,0,'公用',1,'个人',2,'团体',3,'专用') As 组件种类,a.报告元素id,a.组件编码,a.组件名称,a.组件分析,a.组件属性,a.默认分析,a.附加说明 " & vbNewLine & _
            "From 体检报告组件 a "
        Set GetPeisComponet = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle)
    Case 2
        mstrSQL = _
            "Select '组件' As 图标,a.ID,Decode(a.组件种类,0,'公用',1,'个人',2,'团体',3,'专用') As 组件种类,a.报告元素id,b.类名,a.组件编码,a.组件名称,a.组件分析,a.组件属性,a.默认分析,a.附加说明 " & vbNewLine & _
            "From 体检报告组件 a,体检报告元素 b Where a.ID=[1] And a.报告元素id=b.ID"
        Set GetPeisComponet = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, GetCondition(rsCondition, "报告组件id"))
    End Select

    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetPeisReportContent(ByVal bytMode As Byte, ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    On Error GoTo errHand
    '报告内容
    Select Case bytMode
    Case 1
        mstrSQL = _
            "Select a.ID,Decode(a.内容类型,1,'文本','组件') As 图标,Decode(a.内容类型,1,'文本','组件') As 内容类型,a.报告目录id,a.内容序号,a.输出限制,a.内容属性,a.禁止输出,a.内容缩进,a.文本内容,a.文本数据源,a.报告组件id,a.内容属性,a.功能显示,a.新页输出,a.目录级数 " & vbNewLine & _
            "From 体检报告内容 a Where a.报告目录id=[1] Order By a.内容序号 "
        Set GetPeisReportContent = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, GetCondition(rsCondition, "报告目录id"))
    Case 2
        mstrSQL = _
            "Select A.ID,Decode(a.内容类型,1,'文本','组件') As 图标,Decode(a.内容类型,1,'文本','组件') As 内容类型,a.报告目录id,a.内容序号,a.输出限制,b.报告属性,a.内容属性,a.禁止输出,a.内容缩进,a.文本内容,a.文本数据源,a.报告组件id,a.内容属性,a.功能显示,a.新页输出,a.目录级数,b.报告种类 " & vbNewLine & _
            "From 体检报告内容 a,体检报告目录 b Where a.ID=[1] And a.报告目录id=b.ID"
        Set GetPeisReportContent = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, GetCondition(rsCondition, "报告内容id"))
    Case 3              '个人报告打印
        mstrSQL = _
            "Select a.ID,a.报告目录id,a.输出限制,a.禁止输出,a.报告组件id,b.组件属性,c.类名 " & vbNewLine & _
            "From 体检报告内容 a,体检报告组件 b,体检报告元素 c Where a.报告目录id=[1] And a.报告组件id=b.ID And b.报告元素id=c.ID Order By a.内容序号 "
        Set GetPeisReportContent = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, GetCondition(rsCondition, "报告目录id"))
    End Select

    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function Get_Rptfmtno(ByVal 报表编号_In As String, ByVal 报表格式_In As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    On Error GoTo errHand
    mstrSQL = _
        "Select a.序号 from zlRPTFMTs a,zlReports b where a.报表id=b.id and  b.编号=[1] And a.说明=[2]"
    Set Get_Rptfmtno = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, 报表编号_In, 报表格式_In)

    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function CreateCondition() As ADODB.Recordset

    Dim rs As New ADODB.Recordset

    With rs
        .Fields.Append "条件名称", adVarChar, 30
        .Fields.Append "条件结果", adVarChar, 4000
        .Fields.Append "条件类型", adVarChar, 30
        .Open
    End With

    Set CreateCondition = rs

End Function

Private Function SetCondition(ByRef rs As ADODB.Recordset, ByVal strConditionName As String, ByVal strConditionValue As String, Optional ByVal strConditionType As String = "文本") As Boolean

    rs.Filter = ""
    rs.Filter = "条件名称='" & strConditionName & "'"
    If rs.RecordCount = 0 Then rs.AddNew
    rs("条件名称").Value = strConditionName
    rs("条件结果").Value = strConditionValue
    rs("条件类型").Value = strConditionType
    SetCondition = True

End Function

Private Function GetCondition(ByRef rs As ADODB.Recordset, ByVal strConditionName As String) As String
    rs.Filter = ""
    rs.Filter = "条件名称='" & strConditionName & "'"
    If rs.RecordCount > 0 Then
        GetCondition = CStr(rs("条件结果").Value)
    End If
End Function

Private Function GetFileNameExt(ByVal strFile As String, ByRef strName As String, ByRef strExt As String)
    strFile = StrReverse(strFile)
    strName = Mid(strFile, InStr(strFile, ".") + 1)
    strExt = Mid(strFile, 1, InStr(strFile, ".") - 1)
    strExt = StrReverse(strExt)
    strName = StrReverse(strName)
End Function



Private Sub Class_Terminate()
    Set mobjPeisReport = Nothing
    Set mobjReport = Nothing
End Sub
