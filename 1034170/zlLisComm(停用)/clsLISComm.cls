VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private commport As Integer '串口号
Private lngDeviceID As Long '设备ID
Private strDevice As String '设备程序名
Private objDevice As Object '设备解析实例
Private strCommSetting As String '串口设置：速度，校验，数据，停止|握手|接收方式
Private intMicrobe As Integer    '是否是微生物 1= 微生物
Private InMode As Long          ' winsock接收模式

'Private mcnOracle As ADODB.Connection
'Private mcnAccess As New ADODB.Connection

Private mItem() As Variant
Public Event ItemUnknown(ByVal strItems As String)
Public Event Decode(ByVal strReturn As String)
Public Event DecodeErr(ByVal strErr As String)
Public Event AutoCompute(ByVal strReturn As String)

Private mlng允许发送已核收标本 As Long

Public Property Get DecodeProgName() As String
    DecodeProgName = strDevice
End Property

Public Sub InitClsLisComm(ByVal intIndex As Integer)
    '初始化实例，在创建类的实例后执行
    Dim objFile As New Scripting.FileSystemObject
    Dim strAccessPath As String
    Dim rsTmp As New adodb.Recordset, strSQL As String
    Dim lngReturn As Long, lngHost As Long, lngSaveAsID As Long
    
    On Error GoTo DBError
    
'    Set mcnOracle = New ADODB.Connection
'    With mcnOracle
'        If .State = adStateOpen Then .Close
'        .Provider = "MSDataShape"
'        .Open OraConnect
'    End With
    mlng允许发送已核收标本 = g仪器(intIndex).可发已核标本
    
    If g仪器(intIndex).ID > 0 Then
        
        If g仪器(intIndex).类型 = 1 Then
            'TCPIP方式
            
            lngDeviceID = g仪器(intIndex).ID
            lngHost = Val(g仪器(intIndex).主机)
            strCommSetting = lngHost & "|" & g仪器(intIndex).IP
            strCommSetting = strCommSetting & ":" & g仪器(intIndex).IP端口
            
            InMode = Val(g仪器(intIndex).字符模式)
            
            lngSaveAsID = g仪器(intIndex).SaveAsID
        
        Else
            commport = g仪器(intIndex).COM口
            lngDeviceID = g仪器(intIndex).ID
            strCommSetting = g仪器(intIndex).波特率
            strCommSetting = strCommSetting & "," & g仪器(intIndex).校验位
            strCommSetting = strCommSetting & "," & g仪器(intIndex).数据位
            strCommSetting = strCommSetting & "," & g仪器(intIndex).停止位
            strCommSetting = strCommSetting & "|" & g仪器(intIndex).握手
            strCommSetting = strCommSetting & "|" & g仪器(intIndex).字符模式
            
            InMode = g仪器(intIndex).字符模式
            lngSaveAsID = g仪器(intIndex).SaveAsID
        End If
    Else
        Exit Sub
    End If

    
    gstrSQL = "Select 通讯程序名,nvl(微生物,0) as 微生物,使用小组ID From 检验仪器 Where ID=" & lngDeviceID
    OpenRecordset rsTmp, App.ProductName, gcnOracle
    If Not rsTmp.EOF Then strDevice = rsTmp(0): intMicrobe = Nvl(rsTmp(1), 0): lngExeDeptID = Nvl(rsTmp(2), 0)
    
    '-----
    '将lngDeviceID换为 关联仪器 的ID
    If lngSaveAsID > 0 Then lngDeviceID = lngSaveAsID
    '------
    If intMicrobe = 0 Then
        gstrSQL = "Select a.通道编码,a.项目ID,Nvl(a.小数位数,2) As 小数位数 From 诊治所见项目 B,检验仪器项目 A Where a.项目id=b.id and a.仪器ID=[1] order by a.通道编码,b.编码 "
    Else
        gstrSQL = "Select 通道编码,抗生素ID As 项目ID, 2 as 小数位数  From 仪器细菌对照 Where 仪器id = [1] "
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, App.ProductName, lngDeviceID)
    If rsTmp.EOF Then
        ReDim mItem(1, 0) As Variant
        mItem(1, 0) = -1
    Else
        mItem = rsTmp.GetRows
    End If
    
    On Error Resume Next
    Set objDevice = CreateObject(strDevice)
    If objDevice Is Nothing Then
        Call WriteLog("clsLISComm.InitClsLisComm", LOG_错误日志, Err.Number, "解析程序:" & strDevice & "创建失败!" & vbNewLine & Err.Description)
    End If


    Exit Sub
DBError:
    Call WriteLog("clsLISComm.InitClsLisComm", LOG_错误日志, Err.Number, Err.Description)
End Sub

Public Property Get CommSetting() As String
    '串口设置参数
    CommSetting = strCommSetting
End Property

Public Property Get InputMode() As Integer
    '接收方式 0-文本 1－二进制
    InputMode = InMode
End Property

Public Property Get DeviceID() As Long
    '设备ID
    DeviceID = lngDeviceID
End Property

Public Function GetDeviceStartCmd() As String
    '获取设备开始发送命令
    GetDeviceStartCmd = ""
    
    On Error Resume Next
    GetDeviceStartCmd = objDevice.GetStartCmd
    '写日志
    If GetDeviceStartCmd <> "" Then Call WriteLog("clsLISComm.GetDeviceStartCmd", LOG_通讯日志, 0, "GetDeviceStartCmd" & vbNewLine & GetDeviceStartCmd)
End Function

Public Function GetDeviceEndCmd() As String
    '获取设备停止发送命令
    GetDeviceEndCmd = ""
    
    On Error Resume Next
    GetDeviceEndCmd = objDevice.GetEndCmd
    '写日志
    If GetDeviceEndCmd <> "" Then Call WriteLog("clsLISComm.GetDeviceEndCmd", LOG_通讯日志, 0, "GetDeviceEndCmd" & vbNewLine & GetDeviceEndCmd)
    
End Function

Public Function GetDeviceAnswerCmd() As String
    '获取自动应答命令
    GetDeviceAnswerCmd = ""
    On Error Resume Next
    GetDeviceAnswerCmd = objDevice.GetAnswerCmd
    '写日志
    If GetDeviceAnswerCmd <> "" Then Call WriteLog("clsLISComm.GetDeviceAnswerCmd", LOG_通讯日志, 0, "GetDeviceAnswerCmd" & vbNewLine & GetDeviceAnswerCmd)
End Function


Public Function GetSampleInfo(ByVal lngDeviceID As Long, ByVal strSampleDate As String, ByVal strSampleNO As String, ByVal strBarcode As String, Optional strAdviceIDs As String = "", Optional ByVal iType As Integer = 0) As String
    '获取需要向仪器发送的标本信息
    '返回：标本信息。
    '   标本之间以||分隔
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本类型
    '   第4个元素：样本条码
    '   第5个元素：盘号，杯号
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    '  lngDeviceID = 仪器ID
    '  strSampleDate = 日期 格式为 YYYY-MM-DD
    '  strSampleNO = 标本号
    '  strBarcode = 条码
    '  strAdviceIDs =???
    '  iType = 标本类别
    Dim objDevice As Object
    Dim rsTmp As New adodb.Recordset
    Dim lngAdviceID As Long, aAdviceIDs() As String, i As Integer
    Dim bln发送时指定杯号 As Boolean
    
    Dim str标本号 As String, int_急诊 As Integer
    
    On Error GoTo DBErr
    '发往仪器的数据，正常情况下是不核收，病人医嘱发送中执行状态为0 ，而指定了杯号的仪器需要先核收,填写杯号后再发送，所以指定杯号的仪器就不管执行状态。
    
    If mlng允许发送已核收标本 = 0 Then
        bln发送时指定杯号 = False
        gstrSQL = "Select 发送时指定杯号 From 检验仪器 Where Id = [1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "取仪器属性", lngDeviceID)
        Do Until rsTmp.EOF
            bln发送时指定杯号 = Val("" & rsTmp!发送时指定杯号) = 1
            rsTmp.MoveNext
        Loop
    Else
        bln发送时指定杯号 = True
    End If
    
    If Len(strAdviceIDs) = 0 Or Val(strAdviceIDs) = 0 Then
        If Len(Trim(strBarcode)) = 0 Then
            '按标本序号查询
            gstrSQL = "Select TO_CHAR(A.核收时间, 'MM-DD HH24:MI') AS 标本时间,A.标本序号 AS 标本号,F.姓名,D.中文名,D.英文名,C.通道编码,A.标本类型,A.样本条码, A.杯号" & _
                " From 检验标本记录 A,检验普通结果 B,检验仪器项目 C,诊治所见项目 D,病人医嘱记录 E,病人信息 F,检验项目 G,病人医嘱发送 H " & _
                " Where A.ID+0=B.检验标本ID And A.报告结果=B.记录类型 And B.检验项目ID+0=C.项目ID And A.仪器ID=C.仪器ID And C.项目ID=D.ID" & _
                " And A.医嘱ID+0=E.ID And E.病人ID+0=F.病人ID And D.ID=G.诊治项目ID And A.仪器ID=[1]" & _
                " And A.核收时间 BETWEEN [2] AND [3] And E.id=H.医嘱ID " & IIf(bln发送时指定杯号 = True, "", " And H.执行状态 = 0 ") & _
                " And A.标本序号=[4] And G.项目类别<>3 And C.通道编码<>'0' " & IIf(gblnEmerge, " and nvl(a.标本类别,0)  = [5] ", "")
            Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "发送仪器数据", lngDeviceID, CDate(Format(strSampleDate, "yyyy-mm-dd") & " 00:00:00"), _
                CDate(Format(strSampleDate, "yyyy-mm-dd") & " 23:59:59"), Val(strSampleNO), iType)
            Call WriteLog("GetSampleInfo", LOG_通讯日志, 0, "按标本查:" & lngDeviceID & "," & strSampleNO & "," & strSampleDate & "," & iType & "," & mlng允许发送已核收标本)
            
        Else
            '按条码查找
            gstrSQL = "Select TO_CHAR(A.核收时间, 'MM-DD HH24:MI') AS 标本时间,A.标本序号 AS 标本号,F.姓名,D.中文名,D.英文名,C.通道编码,A.标本类型,A.样本条码, A.杯号" & _
                " From 检验标本记录 A,检验普通结果 B,检验仪器项目 C,诊治所见项目 D,病人医嘱记录 E,病人信息 F,检验项目 G,病人医嘱发送 H" & _
                " Where A.ID+0=B.检验标本ID And A.报告结果=B.记录类型 And B.检验项目ID+0=C.项目ID And A.仪器ID=C.仪器ID And C.项目ID=D.ID" & _
                " And A.医嘱ID+0=E.ID And E.病人ID+0=F.病人ID And D.ID=G.诊治项目ID And A.仪器ID=[1]" & _
                " And A.核收时间 BETWEEN [2] AND [3] And E.id=H.医嘱ID " & IIf(bln发送时指定杯号 = True, "", " And H.执行状态 = 0 ") & _
                " And A.样本条码=[5] And G.项目类别<>3 And C.通道编码<>'0' " & IIf(gblnEmerge, " and nvl(a.标本类别,0)  = [6] ", "")
            Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "发送仪器数据", lngDeviceID, CDate(Format(strSampleDate, "yyyy-mm-dd") & " 00:00:00"), _
                CDate(Format(strSampleDate, "yyyy-mm-dd") & " 23:59:59"), Val(strSampleNO), strBarcode, iType)
            If rsTmp.EOF Then
                '查找医嘱  医嘱状态=8 检验医嘱都是临嘱，发送后就是已停止
                gstrSQL = "Select TO_CHAR(F.发送时间, 'MM-DD HH24:MI') AS 标本时间,0 AS 标本号,A.紧急标志," & _
                    "C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)') As 姓名,Y.通道编码,A.标本部位 As 标本类型,F.样本条码,'' as 杯号 " & _
                    "FROM 病人医嘱记录 A," & _
                    "病人信息 C,病人医嘱发送 F,检验报告项目 G,检验项目 I,检验仪器项目 Y " & _
                    "WHERE A.诊疗类别 = 'C' " & _
                    "AND A.病人ID=C.病人ID " & IIf(bln发送时指定杯号 = True, "", " And F.执行状态 = 0 ") & _
                    "AND A.相关id IS NOT NULL " & _
                    "AND A.医嘱状态=8 AND A.ID=F.医嘱id " & _
                    "AND A.诊疗项目id=G.诊疗项目id AND G.细菌ID Is Null AND G.报告项目id=Y.项目id " & _
                    "AND G.报告项目ID=I.诊治项目ID " & _
                    "AND Y.仪器ID+0=[1] " & _
                    "And F.样本条码=[2] " & _
                    IIf(gblnEmerge, "And nvl(A.紧急标志,0) = [3] ", "") & _
                    "And Y.通道编码<>'0' "

                Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "发送仪器数据", lngDeviceID, strBarcode, iType)
            End If
        End If
        GetSampleInfo = ""
        If Not rsTmp.EOF Then
            
            GetSampleInfo = Format(rsTmp("标本时间"), "yyyy-MM-dd HH:mm:ss")
            GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("标本号"), " ")
            GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("姓名"), " ")
            GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("标本类型"), " ") & "|" & iType
            GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("样本条码"), " ") & "|" & IIf(Trim("" & rsTmp("杯号")) = "", " ", Trim("" & rsTmp("杯号"))) & "| | | "
            Do While Not rsTmp.EOF
                GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("通道编码"), " ")
            
                rsTmp.MoveNext
            Loop
        End If
    Else '按医嘱ID查询
        aAdviceIDs = Split(strAdviceIDs, ",")
        GetSampleInfo = ""
        For i = 0 To UBound(aAdviceIDs)
            lngAdviceID = Val(aAdviceIDs(i))
        
            gstrSQL = "Select TO_CHAR(A.核收时间, 'MM-DD HH24:MI') AS 标本时间,A.标本序号 AS 标本号,F.姓名,D.中文名,D.英文名,C.通道编码,A.标本类型,'' As 样本条码, A.杯号" & _
                " From 检验标本记录 A,检验项目分布 B,检验仪器项目 C,诊治所见项目 D,病人医嘱记录 E,病人信息 F,检验项目 G,病人医嘱发送 H " & _
                " Where A.ID=B.标本ID+0 And B.项目ID+0=C.项目ID And A.仪器ID=C.仪器ID And C.项目ID=D.ID" & _
                " And B.医嘱ID=E.ID And E.病人ID+0=F.病人ID And D.ID=G.诊治项目ID And A.仪器ID=[1] And E.id=H.医嘱ID " & IIf(bln发送时指定杯号 = True, "", " And H.执行状态 = 0 ") & _
                " And B.医嘱ID=[2] And G.项目类别<>3 And C.通道编码<>'0' " & IIf(gblnEmerge, " and nvl(a.标本类别,0)  = [3] ", "")
            Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "发送仪器数据", lngDeviceID, lngAdviceID, iType)
            If Not rsTmp.EOF Then
                If Len(GetSampleInfo) = 0 Then
                    GetSampleInfo = Format(rsTmp("标本时间"), "yyyy-MM-dd HH:mm:ss")
                    GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("标本号"), " ")
                    GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("姓名"), " ")
                    GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("标本类型"), " ") & "|" & iType
                    GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("样本条码"), " ") & "|" & IIf(Trim("" & rsTmp("杯号")) = "", " ", Trim("" & rsTmp("杯号"))) & "| | |"
                End If
                Do While Not rsTmp.EOF
                    GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("通道编码"), " ")
                
                    rsTmp.MoveNext
                Loop
            End If
        Next
    End If
    Call WriteLog("getSampleInfo", LOG_通讯日志, 0, GetSampleInfo)
    Exit Function
DBErr:

    GetSampleInfo = ""
    Call WriteLog("clsLISComm.GetSampleInfo", LOG_错误日志, Err.Number, Err.Description & vbCrLf & gstrSQL)
End Function

Public Function SendSample(iSendStep As Integer, strSampleInfo As String, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '解析需要向仪器发送的标本信息
    'iSendStep：当前发送操作的步骤。如果处理完后发送结束，则返回0，否则自动＋1返回。
    'strResponse：当前仪器的响应内容
    'strSampleInfo：发送的标本信息。
    '   标本之间以||分隔
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    Dim aItem() As String
    On Error Resume Next
    SendSample = "": blnSuccess = True
    On Error GoTo DBErr
    
    SendSample = objDevice.SendSample(strSampleInfo, iSendStep, blnSuccess, strResponse, blnUndo, iType)
    WriteLog "clsLISComm.SendSample", LOG_通讯日志, 0, "iSendStep=" & iSendStep & vbNewLine & "SendSample=" & strSampleInfo & vbNewLine & "strResponse=" & strResponse & vbNewLine & "SendSample =" & SendSample
    
    aItem = Split(strSampleInfo, "|")
    If UBound(aItem) > 1 Then
        gstrSQL = "ZL_检验标本记录_传送(" & lngDeviceID & ",To_Date('" & aItem(0) & "','yyyy-MM-dd HH24:mi:ss')," & _
            aItem(1) & "," & IIf(blnSuccess, 1, 0) & "," & iType & ")"
        ExecuteProcedure "传送申请信息", gcnOracle
    End If
    Exit Function
DBErr:
    Call WriteLog("clsLISComm.SendSample", LOG_错误日志, Err.Number, Err.Description)
End Function

Public Function SaveToLocal(ByVal strInput As String) As Long
    '将输入缓冲区的字符存入本地
    Dim strSQL As String, rsTmp As adodb.Recordset
    Dim strContent As String
    
    On Error GoTo SaveError
    SaveToLocal = 0
    
    Exit Function
SaveError:
'    mcnAccess.RollbackTrans
    Call WriteLog("clsLISComm.SaveToLocal", LOG_错误日志, Err.Number, Err.Description)
End Function

Public Function Analyse(ByVal lngDataID As Long, strResult As String, _
    strReserved As String, strCmd As String, Optional ByVal strDevData As String = "", Optional blnGetSample As Boolean) As Boolean
    '解析从串口接收的字符串并保存
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,一组内以回车换行（chr(13)+chr(10)）分隔指标数据和图形数据，元素之间以|分隔
    '   指标数据元素：
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    '   图形数据元素：
    '   第0个元素：图形名称
    '   第1个元素：图形类型 0－直方图、1－闪点图
    '   第2个元素：系统保留
    '   第3个元素：系统保留
    '   第4个元素：系统保留
    '   第5个元素：图形数据
    
    Dim strInput As String
    Dim aCmd() As String
    
    '保存到本地 文本文件中
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Dim lngIndex As Long
    On Error Resume Next
    Analyse = False
    blnGetSample = False
    
    If objDevice Is Nothing Then Exit Function
    
    If lngDataID > 0 Then
        strInput = strDevData
    Else
        strInput = strDevData
    End If
    
    If glngMBDeviceID = lngDeviceID Then
        '酶标仪接口
        Call objDevice.Analyse(strInput, strResult, strReserved, strCmd, gstrMBChannel, glngMBNo, gstrMBPosition)
    Else
        Call objDevice.Analyse(strInput, strResult, strReserved, strCmd)
    End If
    
    Call WriteLog(strDevice & ".Analyse", LOG_通讯日志, 0, "StrInput:" & strInput & vbNewLine & _
                 "strResult:" & strResult & vbNewLine & _
                 "strReserved:" & strReserved & vbNewLine & _
                 "strCmd:" & strCmd & vbNewLine & _
                 "gstrMBChannel:" & gstrMBChannel & vbNewLine & _
                 "glngMBNo:" & glngMBNo & vbNewLine & _
                 "gstrMBPosition:" & gstrMBPosition)
    RaiseEvent Decode(strResult)
    If Err.Number <> 0 Then
        Call WriteLog("clsLisComm.Analyse", LOG_错误日志, Err.Number, Err.Description)
        RaiseEvent DecodeErr(Err.Number & " " & Err.Description)
    End If
    If Len(strCmd) > 0 Then
        aCmd = Split(strCmd, "|")
        If UBound(aCmd) > 0 Then
            strCmd = Mid(strCmd, 3)
            If Val(aCmd(0)) = 1 Then '仪器请求获取标本信息，并非检验结果
                blnGetSample = True: Analyse = True
'                If mcnAccess.State <> adStateClosed Then mcnAccess.CommitTrans
                Exit Function
            End If
        End If
    End If
    
    On Error GoTo DBError
    If Len(strResult) > 0 Then
        If Not objFileSystem.FolderExists(App.Path & "\Apply") Then objFileSystem.CreateFolder (App.Path & "\Apply")
        If Not objFileSystem.FolderExists(App.Path & "\Apply\Decode") Then objFileSystem.CreateFolder (App.Path & "\Apply\Decode")
        lngIndex = 1
        strFileName = App.Path & "\Apply\Decode\" & lngDeviceID & "_" & Format(Now, "yyyyMMddHHmmss") & "_" & lngIndex & ".txt"
        Do
            If Not objFileSystem.FileExists(strFileName) Then Exit Do
            strFileName = App.Path & "\Apply\Decode\" & lngDeviceID & "_" & Format(Now, "yyyyMMddHHmmss") & "_" & lngIndex & ".txt"
            lngIndex = lngIndex + 1
        Loop
        objFileSystem.CreateTextFile (strFileName)
        Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
        If InStr(strResult, vbCrLf) > 0 Then
            objStream.WriteLine Replace(strResult, vbCrLf, "CHR(10) CHR(13)")
        Else
            objStream.WriteLine strResult
        End If
        objStream.Close
        Set objStream = Nothing
    End If
    
    Analyse = True
    Exit Function
DBError:
    'Resume
    Call WriteLog("clsLisComm.Analyse", LOG_错误日志, Err.Number, Err.Description & vbCrLf & gstrSQL)
    RaiseEvent DecodeErr(Err.Number & " " & Err.Description)
End Function

Public Function InDataBase(ByVal strResult As String) As Boolean
    '保存数据到数据库
    Dim aRecord() As String, aItem() As String, aItemInfo() As Variant
    Dim aTmp() As String
    Dim strDate As String, strSampleID As String, strBarcode As String
    Dim strName As String, strSample As String, strSex As String, strBirth As String
    Dim iResultFlag As Integer, strResultRef As String, aResultRef() As String
    Dim i As Long, j As Long
    Dim rsTmp As New adodb.Recordset, strSQL As String, rsContent As adodb.Recordset
    Dim rsRef As New adodb.Recordset
    Dim lngID As Long, strIDs As String '原始数据对应的检验记录ID（可能多个）

    Dim blnAuditing As Boolean '是否审核
    Dim lngItemID As Long '项目ID
    Dim strItemRecords As String
    Dim aNos() As String, iType As Integer '标本号数组
    Dim aQC() As String                    '质控数组
    Dim aMro() As String                   '微生物
    Dim aImage() As String                 '图像数组
    Dim iDec As Integer '小数位数
    Dim blnQryWithSampleNO As Boolean
    Dim blnBeginTrans As Boolean, str未知项 As String
    Dim strStartDate As String
    Dim strEndDate As String
    Dim bln是否区分急诊标本 As Boolean

    Dim strQCList() As String '保存需要计算的内容
    
    On Error GoTo DBError
    ReDim strQCList(0) As String
    InDataBase = False
    
    blnBeginTrans = False
    gcnOracle.BeginTrans
    blnBeginTrans = True
    If Len(strResult) > 0 Then
        aRecord = Split(strResult, "||")
        For i = 0 To UBound(aRecord)
            blnAuditing = False
        
            aTmp = Split(aRecord(i), vbCrLf)
            
            aItem = Split(aTmp(0), "|")
            aQC = Split(aItem(4), "^")              '标记质控
            If UBound(aItem) >= 4 Then
                '有效的报告组
                aNos = Split(aItem(1), "^") '标本号格式：标本号^标本类别^SampleID（0：常规，1：急诊）
                If UBound(aNos) = 0 Then
                    '没有标本类别，则按常规标本处理
                    strDate = Trim(aItem(0)): strSampleID = IIf(aQC(0) = "1", aNos(0), Val(aNos(0))): iType = 0: strBarcode = ""
                Else
                    strDate = Trim(aItem(0)): strSampleID = IIf(aQC(0) = "1", aNos(0), Val(aNos(0))): iType = Val(aNos(1)): strBarcode = ""
                    If UBound(aNos) > 1 Then
                        strBarcode = Trim(aNos(2))
                    End If
                End If
                '单独处理标本生成规则（按时间）
                strStartDate = GetDateTime(mMakeNoRule, 1, strDate)
                strEndDate = GetDateTime(mMakeNoRule, 2, strDate)
                
                strName = Trim(aItem(2)): strSample = Trim(aItem(3))
                '判断是否无主标本
                If Len(Trim(strBarcode)) = 0 Then
                    '按标本号查
                    blnQryWithSampleNO = True
                Else
                    '按条码查询
                    gstrSQL = "Select a.*,Decode(A.性别,Null,0,'男',1,'女',2,0) As 性别A,to_char(c.出生日期,'yyyy-mm-dd') As 出生日期A From 检验标本记录 a,病人医嘱记录 b,病人信息 c " & _
                        " Where a.医嘱id=b.id(+) And b.病人id=c.病人id(+)" & _
                        " And a.核收时间 Between [1] And [2]" & _
                        " And a.仪器ID=[3] And a.样本条码=[6]"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "查询标本记录", CDate(strStartDate), _
                        CDate(strEndDate), lngDeviceID, strSampleID, iType, strBarcode)
                    If Not rsTmp.EOF Then
                        blnQryWithSampleNO = False
                    Else
                        '检验是否已有标本
                        gstrSQL = "Select a.*,Decode(A.性别,Null,0,'男',1,'女',2,0) As 性别A,to_char(c.出生日期,'yyyy-mm-dd') As 出生日期A From 检验标本记录 a,病人医嘱记录 b,病人信息 c " & _
                        " Where a.医嘱id=b.id(+) And b.病人id=c.病人id(+)" & _
                        " And a.核收时间 Between [1] And [2]" & _
                        " And a.仪器ID=[3] And a.标本序号=[4] " & IIf(gblnEmerge, " And Nvl(a.标本类别,0)=[5]", "")
                        Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "查询标本记录", CDate(Format(strDate, "yyyy-MM-dd") & " 00:00:00"), _
                            CDate(Format(strDate, "yyyy-MM-dd") & " 23:59:59"), lngDeviceID, strSampleID, iType, strBarcode)
                        If rsTmp.EOF = True Then
                            '根据条码生成标本
                            Call CreateSample(lngDeviceID, strBarcode, strSampleID, CDate(strDate), iType)
                            blnQryWithSampleNO = True
                        Else
                            If Val(Nvl(rsTmp("医嘱id"))) = 0 Then
                                '标本为无主时也生成
                                Call CreateSample(lngDeviceID, strBarcode, strSampleID, CDate(strDate), iType)
                                blnQryWithSampleNO = True
                            End If
                        End If
                    End If
                End If
                If blnQryWithSampleNO Then
                    gstrSQL = "Select a.*,Decode(A.性别,Null,0,'男',1,'女',2,0) As 性别A,to_char(c.出生日期,'yyyy-mm-dd') As 出生日期A From 检验标本记录 a,病人医嘱记录 b,病人信息 c " & _
                        " Where a.医嘱id=b.id(+) And b.病人id=c.病人id(+)" & _
                        " And a.核收时间 Between [1] And [2]" & _
                        " And a.仪器ID=[3] And a.标本序号=[4] " & IIf(gblnEmerge, " And Nvl(a.标本类别,0)=[5]", "")
                    Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "查询标本记录", CDate(strStartDate), _
                        CDate(strEndDate), lngDeviceID, strSampleID, iType, strBarcode)
                
                End If
                
                If rsTmp.EOF Then
                    '无主标本增加临时标本记录
                    strSex = 0
                    strBirth = ""
                    
                    lngID = gobjDatabase.GetNextId("检验标本记录")
                    
                    gstrSQL = "ZL_检验标本记录_INSERT(" & lngID & ",NULL,'" & _
                        strSampleID & "',NULL,NULL," & lngDeviceID & ",NULL," & _
                        "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),NULL," & _
                        "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & strSample & "'," & _
                        "Null,To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & strName & "','0'," & lngExeDeptID & "," & iType & "," & intMicrobe & ")"
                    
                    ExecuteProcedure "插入检验临时记录", gcnOracle
                    
                Else
                    strSex = Nvl(rsTmp("性别A"), 0)
                    strBirth = Nvl(rsTmp("出生日期A"))
                    If intMicrobe = 0 Then
                        strSample = Nvl(rsTmp("标本类型"))
                    End If
                    lngID = rsTmp("ID")
                    blnAuditing = Not IsNull(rsTmp("初审人"))
                    If blnAuditing = False Then
                        blnAuditing = Not IsNull(rsTmp("审核人"))
                    End If
                End If
                
                
                
                If Not blnAuditing Then
                    If InStr(strIDs, "," & lngID) = 0 Then strIDs = strIDs & "," & lngID
                    '处理检验项目
                    strItemRecords = ""
                    str未知项 = ""
                    For j = 5 To UBound(aItem) Step 2
                        '根据通道号修改相应项目结果，未找到的则直接增加（根据通道号找不到项目的暂不处理）
                        '根据通道号找项目
                        lngItemID = GetItemID(aItem(j), iDec)
                        If lngItemID > 0 Then
                            
                            gstrSQL = "select 项目id from 检验仪器项目 where 仪器id = [1] and 糖耐量项目 = -1 and 项目id = [2] "
                            Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "糖耐量", lngDeviceID, lngItemID)
                            If rsTmp.EOF = False Then
                                '仪器有糖耐量项目时的处理
                                If strBarcode <> "" Then
                                    '有条码时的处理 ,根据通道码
                                    gstrSQL = "Select d.项目id" & vbNewLine & _
                                            "From 病人医嘱记录 A, 病人医嘱发送 B, 检验报告项目 C, 检验仪器项目 D" & vbNewLine & _
                                            "Where A.ID = B.医嘱id And B.样本条码 = [2] And A.诊疗项目id = C.诊疗项目id And C.报告项目id = D.项目id" & vbNewLine & _
                                            "      And D.仪器id = [1] And D.通道编码 =[3] And D.糖耐量项目 = -1"
                                    Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "仪器糖耐量", lngDeviceID, strBarcode, CStr(aItem(j)))
                                    If rsTmp.EOF = False Then
                                        strItemRecords = strItemRecords & "|" & Nvl(rsTmp("项目ID")) & "^" & aItem(j + 1)
                                    Else
                                        strItemRecords = strItemRecords & "|" & lngItemID & "^" & aItem(j + 1)
                                    End If
                                Else
                                    '没有条码时的处理
                                    gstrSQL = "Select B.项目id" & vbNewLine & _
                                            " From 检验普通结果 A, 检验仪器项目 B" & vbNewLine & _
                                            " Where A.检验项目id = B.项目id And B.仪器id = [1] And B.糖耐量项目 = -1 And B.通道编码=[3]  And A.检验标本id = [2] "
                                    Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "仪器糖耐量", lngDeviceID, lngID, CStr(aItem(j)))
                                    If rsTmp.EOF = False And rsTmp.RecordCount = 1 Then
                                        strItemRecords = strItemRecords & "|" & Nvl(rsTmp("项目ID")) & "^" & aItem(j + 1)
                                    Else
                                        strItemRecords = strItemRecords & "|" & lngItemID & "^" & aItem(j + 1)
                                    End If
                                End If
                            Else
                                '仪器没有糖耐量项目时的处理
                                strItemRecords = strItemRecords & "|" & lngItemID & "^" & aItem(j + 1)
                            End If
                        Else
'                            strSql = "Delete From 未知项目 Where 标本记录ID=" & lngID & _
'                                " And 项目标识='" & aItem(j) & "'"
'                            mcnAccess.Execute strSql
'                            If j + 1 <= UBound(aItem) Then
'                                strSql = "Insert Into 未知项目(标本记录ID,项目标识,项目值)" & _
'                                    " Values(" & lngID & ",'" & aItem(j) & "','" & aItem(j + 1) & "')"
'                            End If
                            If str未知项 = "" Then str未知项 = "标本号     项目标识     项目值" & vbNewLine
                            str未知项 = str未知项 & strSampleID & Space(11 - Len(strSampleID)) & _
                            aItem(j) & Space(30 - Len(aItem(j))) & _
                            aItem(j + 1) & vbNewLine
'                            mcnAccess.Execute strSql
                        End If
                    Next
                    If str未知项 <> "" Then Call WriteLog("clsLISComm.Analyse", LOG_未知项, 0, str未知项)
                    
                    RaiseEvent ItemUnknown(str未知项)
                    
                    If Len(strItemRecords) > 0 Then
                        strItemRecords = Mid(strItemRecords, 2)
                        
                        gstrSQL = "ZL_检验普通结果_BATCHUPDATE(" & lngID & "," & _
                            lngDeviceID & ",'" & strSample & "'," & strSex & "," & _
                            IIf(strBirth = "", "Null", "To_Date('" & strBirth & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                            strItemRecords & "'," & intMicrobe & ")"
                        ExecuteProcedure "检验结果报告", gcnOracle
                        
                        '重新计算计算项目
'                        gstrSQL = "Zl_重新计算结果_Cale(" & lngID & ")"
'                        ExecuteProcedure "重新计算结果", gcnOracle
                        
                        '保存为质控
                        If aQC(0) = 1 Then
                            Dim date当前日期 As Date, lngQCID As Long, str标本号 As String
                            Dim var标本号 As Variant, iCoutn As Integer
                            lngQCID = 0
                            date当前日期 = gobjDatabase.Currentdate
'                            gstrSQL = "Select ID From 检验质控品 Where 仪器id = [1] And 标本号 = [2]"
                            gstrSQL = "Select ID,标本号 From 检验质控品 Where [2] between 开始日期 and 结束日期 And 仪器id = [1] "
                            Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lngDeviceID, date当前日期)
                            
                            Do Until rsTmp.EOF Or lngQCID <> 0
                                str标本号 = "" & rsTmp.Fields("标本号")
                                If InStr(str标本号, ",") > 0 Then
                                    var标本号 = Split(str标本号, ",")
                                    For iCoutn = 0 To UBound(var标本号)
                                        If var标本号(iCoutn) Like "*-*" Then
                                            If strSampleID >= Val(Split(var标本号(iCoutn), "-")(0)) And strSampleID <= Val(Split(var标本号(iCoutn), "-")(1)) Then
                                                lngQCID = rsTmp.Fields("ID")
                                            End If
                                        Else
                                            If var标本号(iCoutn) = strSampleID Then
                                                lngQCID = rsTmp.Fields("ID")
                                            End If
                                        End If
                                    Next
                                ElseIf str标本号 Like "*-*" Then
                                    If strSampleID >= Val(Split(str标本号, "-")(0)) And strSampleID <= Val(Split(str标本号, "-")(1)) Then
                                        lngQCID = rsTmp.Fields("ID")
                                    End If
                                Else
                                    If strSampleID = str标本号 Then
                                        lngQCID = rsTmp.Fields("ID")
                                    End If
                                End If
                                
                                rsTmp.MoveNext
                            Loop
                            
                            If lngQCID > 0 Then
                                gstrSQL = "ZL_检验质控记录_EDIT(1," & lngID & "," & lngQCID & ")"
                                ExecuteProcedure "保存为质控品", gcnOracle
                                
                                '要自动计算
                                If strQCList(UBound(strQCList)) <> "" Then ReDim Preserve strQCList(UBound(strQCList) + 1)
                                strQCList(UBound(strQCList)) = Format(CDate(strDate), "yyyy-MM-dd") & "," & CStr(lngQCID)
                            End If
                        End If
                    End If
                End If
        
                If UBound(aTmp) > 0 Then
                    If Trim(aTmp(1)) <> "" Then
                        '处理图形数据
                        Call WriteLog("SaveImg", LOG_通讯日志, 0, "开始时间:" & Format(Now(), "yyyy-MM-dd HH:mm:ss"))
                        Call SaveImg(lngID, aTmp(1))
                        Call WriteLog("SaveImg", LOG_通讯日志, 0, "结束时间:" & Format(Now(), "yyyy-MM-dd HH:mm:ss"))
                    End If
                End If
            End If
        Next
    End If
   
    gcnOracle.CommitTrans
    blnBeginTrans = False
    '计算质控

    InDataBase = True
    For i = LBound(strQCList) To UBound(strQCList)
        If InStr(strQCList(i), ",") > 0 Then
            Call AutoQCCompute(CDate(Split(strQCList(i), ",")(0)), Split(strQCList(i), ",")(1))
        End If
    Next
    Exit Function
DBError:
    
    If blnBeginTrans Then gcnOracle.RollbackTrans
    Call WriteLog("clsLisComm.InDataBase", LOG_错误日志, Err.Number, Err.Description & vbCrLf & gstrSQL)
    RaiseEvent DecodeErr(Err.Number & " " & Err.Description)
    
End Function
Private Sub AutoQCCompute(ByVal date日期 As Date, ByVal str质控品 As String)

    '自动计算质控标本
    ' date日期 :质控计算日期
    ' str质控品 :质控品
    Dim rsTemp As adodb.Recordset, rsTmp As adodb.Recordset, strReturn As String
    On Error GoTo errH
    gstrSQL = "Select Distinct B.项目id, C.编码, C.中文名, C.英文名" & vbNewLine & _
              " From 检验质控品 A, 检验质控品项目 B, 诊治所见项目 C" & vbNewLine & _
              " Where A.ID = B.质控品id And B.项目id = C.ID And A.仪器id = [1] "
        
    Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "LisComm自动计算", lngDeviceID)
    Do Until rsTmp.EOF
        '计算一段时间
            gstrSQL = "Select Zl_检验质控记录_Compute(" & lngDeviceID & ", " & rsTmp("项目ID") & ", To_Date('" & Format(date日期, "yyyy-mm-dd") & "','yyyy-mm-dd'), '" & str质控品 & "') From Dual"
            Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, "LisComm自动计算")

            If rsTemp.RecordCount <= 0 Then strReturn = strReturn & Format(date日期, "yyyy-mm-dd") & " " & Nvl(rsTmp("中文名")) & "(" & Nvl(rsTmp("英文名")) & ")  计算过程调用错误！" & vbCrLf
            If InStr(rsTemp.Fields(0).Value, "出现失控！") > 0 Then
                strReturn = strReturn & Format(date日期, "yyyy-mm-dd") & " " & Nvl(rsTmp("中文名")) & "(" & Nvl(rsTmp("英文名")) & ")" & rsTemp.Fields(0).Value & vbCrLf

            ElseIf InStr(rsTemp.Fields(0).Value, "计算完成！") <= 0 Then
                If InStr(rsTemp.Fields(0).Value, "按规则未发现警告和失控！") <= 0 Then
                strReturn = strReturn & Format(date日期, "yyyy-mm-dd") & " " & Nvl(rsTmp("中文名")) & "(" & Nvl(rsTmp("英文名")) & ")" & rsTemp.Fields(0).Value & vbCrLf
                End If
            End If
        rsTmp.MoveNext
    Loop
    If Trim(strReturn) <> "" Then
       RaiseEvent AutoCompute(strReturn)
    End If
    Exit Sub
errH:
   WriteLog "AutoQCCompute", LOG_错误日志, Err.Number, Err.Description
End Sub
Private Sub Class_Terminate()
'    If Not mcnAccess Is Nothing Then
'        If mcnAccess.State <> adStateClosed Then mcnAccess.Close
'    End If
''    If Not mcnOracle Is Nothing Then
''        If mcnOracle.State <> adStateClosed Then mcnOracle.Close
''    End If
''    Set mcnOracle = Nothing
'    Set mcnAccess = Nothing
End Sub

Private Function GetItemID(ByVal strChannel As String, ByRef iDec As Integer) As Long
    'iDec:小数位数
    Dim i As Integer
    For i = 0 To UBound(mItem, 2)
        If Trim(Replace(Replace(UCase(strChannel), Chr(10), ""), Chr(13), "")) = _
           Replace(Replace((UCase(mItem(0, i))), Chr(10), ""), Chr(13), "") Then Exit For
    Next
    If i > UBound(mItem, 2) Then
        GetItemID = -1
        iDec = 2
    Else
        GetItemID = CLng(mItem(1, i))
        iDec = Val(mItem(2, i))
    End If
End Function

Public Sub test()
    CreateSample 149, 120001037818#, "H343343", Now, 0
End Sub




