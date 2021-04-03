Attribute VB_Name = "mdlLISComm"
Option Explicit

'Public gcnOracle As ADODB.Connection    '公共数据库连接
Public gstrSQL As String

'Public gstrSysName As String                '系统名称

Public lngExeDeptID As Long '执行科室
Public ParentWnd As Object
Public blnDataReceived As Boolean
'------任务栏图标处理
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_ACTIVATE = &H6
Public Const WM_KEYDOWN = &H100
Public Const WM_PAINT = &HF

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

'Public Const GWL_EXSTYLE = (-20)
'Public Const WinStyle = &H40000
Public Const SW_RESTORE = 9
Public Const GWL_WNDPROC = -4

'酶标仪参数
Public glngMBDeviceID As Long, gstrMBChannel As String, glngMBNo As Long, gstrMBPosition As String

Private mItem() As Variant

Public Const LOG_错误日志 = 0
Public Const LOG_通讯日志 = 1
Public Const LOG_未知项 = 2

Public pLast错误日志 As String '上次错误信息,用于避免输出重复的日志
Public pLast通讯日志 As String
Public mMakeNoRule As String    '标本序号生成时间规则

Public gblnFromDB As Boolean ' 是否是从数据库读取参数.

Public gobjFSO As New Scripting.FileSystemObject    'FSO对象
Public mclsUnzip As New cUnzip
Public mclsZip As New cZip

Public Sub SavePortsSetting()
'功能：保存连接检验仪器的串口设置
    Dim i As Integer
    Dim strSet As String
    Dim aPorts As Variant
    On Error GoTo errH
    
    strSet = ""
    If gblnFromDB Then
        '清空原来的设置
        Call gobjDatabase.SetPara("本机连接仪器", "", glngSys, 1208)
        For i = LBound(g仪器) To UBound(g仪器)
            '仪器id , 类型, COM口, 波特率, 数据位, 校验位, 停止位, 握手, TCPIP端口, IP地址, 字符模式, 另存为的仪器ID, 主机,自动应答,可发已核标本
            If g仪器(i).ID > 0 Then
                strSet = strSet & ";" & g仪器(i).ID & "," & g仪器(i).类型 & "," & g仪器(i).COM口 & "," & g仪器(i).波特率 & _
                   "," & g仪器(i).数据位 & "," & g仪器(i).校验位 & "," & g仪器(i).停止位 & "," & g仪器(i).握手 & _
                   "," & g仪器(i).IP端口 & "," & g仪器(i).IP & "," & g仪器(i).字符模式 & "," & g仪器(i).SaveAsID & "," & g仪器(i).主机 & _
                   "," & g仪器(i).自动应答 & "," & g仪器(i).可发已核标本
            End If
        Next
        If strSet <> "" Then
            Call gobjDatabase.SetPara("本机连接仪器", strSet, glngSys, 1208)
        End If
    Else
        'DeleteSetting "ZLSOFT", "公共模块", "ZlLISSrv"
        Err = 0: On Error Resume Next
        aPorts = GetAllSettings("ZLSOFT", "公共模块\ZlLISSrv")
        On Error GoTo errH
        If IsEmpty(aPorts) Then
            ReDim aPorts(8, 0)
            For i = 0 To 7
                aPorts(i, 0) = "COM" & i + 1
            Next
        End If
        Err = 0: On Error Resume Next
        For i = LBound(aPorts) To UBound(aPorts)
            DeleteSetting "ZLSOFT", "公共模块\ZLLISSrv", aPorts(i, 0)
            DeleteSetting "ZLSOFT", "公共模块\ZLLISSrv\" & aPorts(i, 0)
        Next
        On Error GoTo errH
        For i = LBound(g仪器) To UBound(g仪器)
            If g仪器(i).类型 = 1 Then
                'TCP
                If g仪器(i).ID > 0 Then
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv", "IP" & g仪器(i).ID, "")
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "Device", g仪器(i).ID)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "Enabled", g仪器(i).类型)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "Host", g仪器(i).主机)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "InMode", g仪器(i).字符模式)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "IP", g仪器(i).IP)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "Port", g仪器(i).IP端口)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "SaveAs", g仪器(i).SaveAsID)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "Auto", g仪器(i).自动应答)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "blnSend", g仪器(i).可发已核标本)
                End If
            Else
                If g仪器(i).COM口 > 0 And g仪器(i).ID > 0 Then
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv", "COM" & g仪器(i).COM口, "")
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "Device", g仪器(i).ID)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "Speed", g仪器(i).波特率)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "DataBit", g仪器(i).数据位)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "Parity", g仪器(i).校验位)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "StopBit", g仪器(i).停止位)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "HandShaking", g仪器(i).握手)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "InputMode", g仪器(i).字符模式)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "SaveAs", g仪器(i).SaveAsID)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "Auto", g仪器(i).自动应答)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "blnSend", g仪器(i).可发已核标本)
                End If
            End If
        Next
    End If
    Exit Sub
errH:
    MsgBox Err.Description

End Sub

Public Function GetConnectDevs() As Variant
'功能：获取系统连接的检验仪器
    Dim aSettings() As Variant
    Dim aPorts As Variant, i As Integer, PortIndex As Integer
    Dim lngDeviceID As Long, rsTmp As New adodb.Recordset, rsTmp1 As New adodb.Recordset
    Dim strConnType As String  '连接类型
    Dim strIP As String, strPort As String 'ip 和 Port
    Dim varIPSet As Variant 'IP的设置
    Dim lngSaveAsID As Long '另存为的仪器ID
    Dim strSaveAsName As String
    
    aSettings = Array()
    
    Err = 0: On Error Resume Next
    aPorts = GetAllSettings("ZLSOFT", "公共模块\ZlLISSrv")
    On Error GoTo errH
    If IsEmpty(aPorts) Then
        ReDim aPorts(8, 0)
        For i = 0 To 7
            aPorts(i, 0) = "COM" & i + 1
        Next
    End If
   
    If Not IsEmpty(aPorts) Then
        
        ReDim g仪器(UBound(aPorts))
        
        For i = LBound(g仪器) To UBound(g仪器)
            g仪器(i).ID = 0
            g仪器(i).IP = "127.0.0.1"
            g仪器(i).IP端口 = 6666
            g仪器(i).SaveAsID = 0
            g仪器(i).波特率 = 9600
            g仪器(i).类型 = 1
            g仪器(i).COM口 = 0
            g仪器(i).数据位 = 8
            g仪器(i).停止位 = 1
            g仪器(i).握手 = 0
            g仪器(i).校验位 = "N"
            g仪器(i).字符模式 = 0
            g仪器(i).主机 = 0
            g仪器(i).自动应答 = "0"
            g仪器(i).可发已核标本 = 1
        Next
        
        For i = LBound(aPorts) To UBound(aPorts)
            
            strIP = "": strPort = ""
            lngSaveAsID = 0
            strSaveAsName = ""
            
            lngSaveAsID = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "SaveAs", 0))
            If lngSaveAsID > 0 Then
                Set rsTmp1 = gobjDatabase.OpenSQLRecord("Select 名称 From 检验仪器 where ID=[1]", "取另存检验仪器名", lngSaveAsID)
                Do Until rsTmp1.EOF
                    strSaveAsName = "" & rsTmp1!名称
                    rsTmp1.MoveNext
                Loop
            End If
            
            strConnType = aPorts(i, 0)

            If strConnType Like "IP*" Then
                'TCPIP连接
                g仪器(i).类型 = 1
                lngDeviceID = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Device", 0))
                
                If lngDeviceID > 0 Then

                    If rsTmp.State <> adStateClosed Then rsTmp.Close
                    gstrSQL = "Select * From 检验仪器 Where ID=" & lngDeviceID
                    OpenRecordset rsTmp, App.ProductName
                    If Not rsTmp.EOF Then

                        If Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Enabled", 0)) = 1 Then
                            '启用了IP方式,检查IP和端口是否合法
                            strIP = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "IP", "127.0.0.1")
                            strPort = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Port", 6666)
                            g仪器(i).IP = strIP
                            g仪器(i).IP端口 = Val(strPort)
                            g仪器(i).主机 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Host", 0))
                            
                            g仪器(i).自动应答 = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Auto", "0")
                            g仪器(i).可发已核标本 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "blnSend", "1"))
                            If Not ValidateIP(strIP) And Not ValidatePort(strPort) Then

                                If UBound(aSettings) = -1 Then
                                    ReDim aSettings(2, 0) As Variant
                                Else
                                    ReDim Preserve aSettings(2, UBound(aSettings, 2) + 1) As Variant
                                End If

                                aSettings(0, UBound(aSettings, 2)) = strIP & ":" & strPort
                                aSettings(1, UBound(aSettings, 2)) = "IP " & strIP & " " & rsTmp("名称") & IIf(strSaveAsName = "", "", " -> " & strSaveAsName)
                                aSettings(2, UBound(aSettings, 2)) = lngDeviceID
                            End If

                        End If
                    End If
                End If
            ElseIf strConnType Like "COM*" Then
                'COM连接
                PortIndex = Val(Mid(aPorts(i, 0), 4)) - 1
                lngDeviceID = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Device", 0))
                g仪器(i).类型 = 0
                g仪器(i).COM口 = Val(PortIndex + 1)
                If lngDeviceID > 0 Then
                    If rsTmp.State <> adStateClosed Then rsTmp.Close
                    gstrSQL = "Select * From 检验仪器 Where ID=" & lngDeviceID
                    OpenRecordset rsTmp, App.ProductName
                    If Not rsTmp.EOF Then
                        If UBound(aSettings) = -1 Then
                            ReDim aSettings(2, 0) As Variant
                        Else
                            ReDim Preserve aSettings(2, UBound(aSettings, 2) + 1) As Variant
                        End If
                        aSettings(0, UBound(aSettings, 2)) = PortIndex
                        aSettings(1, UBound(aSettings, 2)) = "COM" & PortIndex + 1 & " " & rsTmp("名称") & IIf(strSaveAsName = "", "", " -> " & strSaveAsName)
                        aSettings(2, UBound(aSettings, 2)) = lngDeviceID
                    End If
                
                    With g仪器(i)
                        .ID = lngDeviceID
                        .波特率 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Speed", "9600"))
                        .数据位 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "DataBit", "8"))
                        .停止位 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "StopBit", "1"))
                        .校验位 = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Parity", "n")
                        .握手 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & aPorts(i, 0), "HandShaking", "0"))
                        .字符模式 = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "InputMode", "0")
                        .SaveAsID = lngSaveAsID
                        .自动应答 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Auto", "0"))
                        .可发已核标本 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "blnSend", "1"))
                    End With
                End If
            End If
        Next
    End If
    
    If UBound(aSettings) > -1 Then GetConnectDevs = aSettings
    Exit Function
errH:
    MsgBox Err.Description

End Function

Public Function GetDevices() As adodb.Recordset
'功能：获取所有检验仪器
    On Error GoTo DBError
    Set GetDevices = Nothing
    gstrSQL = "Select * From 检验仪器"
    Set GetDevices = gobjDatabase.OpenSQLRecord(gstrSQL, "仪器数据接收")
    Exit Function
DBError:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function GetComboxIndex(objCbo As ComboBox, ByVal SeekValue As Long) As Long
    Dim i As Long
    
    For i = 0 To objCbo.ListCount - 1
        If objCbo.ItemData(i) = SeekValue Then Exit For
    Next
    If i > objCbo.ListCount - 1 Then i = 0
    GetComboxIndex = i
End Function

Public Sub OpenRecordset(rsTemp As adodb.Recordset, ByVal strFormCaption As String, Optional cnOracle As adodb.Connection = Nothing)
'功能：打开记录。同时保存SQL语句
    If rsTemp.State = adStateOpen Then rsTemp.Close
    If cnOracle Is Nothing Then Set cnOracle = gcnOracle
    
    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, gstrSQL)
    rsTemp.Open gstrSQL, cnOracle, adOpenStatic, adLockReadOnly
    Call gobjComLib.SQLTest
End Sub

Public Sub ExecuteProcedure(ByVal strFormCaption As String, Optional cnOracle As adodb.Connection = Nothing)
'功能：执行过程式的SQL语句
    If cnOracle Is Nothing Then Set cnOracle = gcnOracle
    
    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, gstrSQL)
    cnOracle.Execute gstrSQL, , adCmdStoredProc
    Call gobjComLib.SQLTest
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub WriteLog(ByVal ModuleName As String, ByVal ErrorType As Integer, ByVal ErrorNum As Long, ByVal ErrorDesc As String)
    'Module:模块或函数名称
    'ErrorType:日志类型
    'errorNum:错误号或日志编号
    'errorDesc:错误信息或日志信息
    Dim strSQL As String
    
    Call WriteTxtLog(ErrorType, ModuleName, IIf(ErrorNum = 0, "", " ") & ErrorDesc)
    
End Sub

Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "")
    
    '功能：在任务栏上增加一个图标
    
    Dim t As NOTIFYICONDATA
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '事件发生的载体，为了不与其它鼠标事件相冲突，所以单独放一个控件
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = IIf(Len(strTip) = 0, "仪器数据接收", strTip) & Chr$(0)

    Shell_NotifyIcon NIM_ADD, t
    
End Sub

Public Sub ModifyIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "", Optional ByVal blnMessage As Boolean = True)
    
    '功能：在任务栏上增加一个图标
    
    Dim t As NOTIFYICONDATA
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '事件发生的载体，为了不与其它鼠标事件相冲突，所以单独放一个控件
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = IIf(blnMessage, WM_MOUSEMOVE, 0)
    t.hIcon = stdIcon
    t.szTip = IIf(Len(strTip) = 0, "仪器数据接收", strTip) & Chr$(0)

    Shell_NotifyIcon NIM_MODIFY, t
    
End Sub

Public Sub RemoveIcon(ByVal lngHwnd As Long)
    
    '功能：从任务栏上删除图标
    
    Dim t As NOTIFYICONDATA
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '事件发生的载体
    t.uId = 1&
    
    Shell_NotifyIcon NIM_DELETE, t
End Sub

Public Sub ResultFromFile(ByVal strFile As String, ByVal lngDeviceID As Long, ByVal strSampleNO As String, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31"))
'从文件获取检验结果
'   strFile：包含路径的文件名
'   lngDeviceID：检验设备ID
'   strSampleNO：标本号。为-1表示选取所有时间范围内的标本
'   dtStart：开始时间。如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
'   dtEnd：结束时间。只有当选取多个标本（strSampleNO=-1）时，该参数有效。如果不指定该参数值，则时间范围为>=dtStart。
    Dim rsTmp As New adodb.Recordset
    Dim strDevice As String
    Dim objDevice As Object, strInput As String
    Dim aRecord() As String, aItem() As String, aItemInfo() As Variant
    Dim strDate As String, strSampleID As String '2007-08-16 改为字符型
    Dim strName As String, strSample As String, strSex As String, strBirth As String
    Dim iResultFlag As Integer, strResultRef As String, aResultRef() As String
    Dim i As Long, j As Long
    Dim strSQL As String, rsContent As adodb.Recordset
    Dim rsRef As New adodb.Recordset
    Dim lngID As Long
    Dim blnAuditing As Boolean '是否审核
    Dim lngItemID As Long '项目ID
    Dim strItemRecords As String
    Dim aNos() As String, iType As Integer '标本号数组
    Dim blnBeginTrans As Boolean, str未知项 As String
    Dim intMicrobe As Integer   '微生物 =1 表示微生物
    Dim strStartDate As String
    Dim strEndDate As String
    Dim strBarcode As String
    Dim blnQryWithSampleNO As Boolean

    Dim aTmp() As String '分隔图形数据
    
    If Len(Trim(strFile)) = 0 Then Exit Sub
    
    gstrSQL = "Select 通讯程序名,nvl(微生物,0) as 微生物 From 检验仪器 Where ID=" & lngDeviceID
    OpenRecordset rsTmp, App.ProductName, gcnOracle
    If Not rsTmp.EOF Then strDevice = rsTmp(0): intMicrobe = Nvl(rsTmp(1), 0)
    
    If intMicrobe = 0 Then
        gstrSQL = "Select 通道编码,项目ID,Nvl(小数位数,2) As 小数位数 From 检验仪器项目 Where 仪器ID=[1]"
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
    If objDevice Is Nothing Then Call WriteLog("ResultFromFile", LOG_错误日志, Err.Number, "解析程序:" & strDevice & "创建失败!" & vbNewLine & Err.Description)
    On Error GoTo DBError
    
    blnBeginTrans = False
    gcnOracle.BeginTrans
    blnBeginTrans = True
    Call WriteLog(strDevice & ".ResultFromFile", LOG_通讯日志, 0, "strFile:" & strFile & vbNewLine & "strSampleNO:" & strSampleNO & vbNewLine & "dtStart:" & CStr(dtStart) & vbNewLine & "dtEnd:" & CStr(dtEnd))
    aRecord = objDevice.ResultFromFile(strFile, strSampleNO, dtStart, dtEnd)
    'aRecord：返回的检验结果数组(各解析程序必须按以下标准组织结果)
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    
    '有返回结果
    If UBound(aRecord) > -1 Then
        
        
        For i = 0 To UBound(aRecord)
            Call WriteLog("mdlLISComm.ResultFromFile", LOG_通讯日志, 0, "记录" & i & ":" & aRecord(i))
            blnAuditing = False
            
            If InStr(aRecord(i), "|") > 0 Then
                aTmp = Split(aRecord(i), vbCrLf)
                aItem = Split(aTmp(0), "|")
                If UBound(aItem) > 4 Then
                    '有效的报告组
                    aNos = Split(aItem(1), "^") '标本号格式：标本号^标本类别（0：常规，1：急诊）
                    If UBound(aNos) = 0 Then
                        '没有标本类别，则按常规标本处理
                        strDate = Trim(aItem(0)): strSampleID = Val(aNos(0)): iType = 0: strBarcode = ""
                    Else
                        strDate = Trim(aItem(0)): strSampleID = Val(aNos(0)): iType = Val(aNos(1)): strBarcode = ""
                        If UBound(aNos) > 1 Then
                            strBarcode = Trim(aNos(2))
                        End If
                    End If
                    '单独处理标本生成规则（按时间）
                    strStartDate = GetDateTime(mMakeNoRule, 1, strDate)
                    strEndDate = GetDateTime(mMakeNoRule, 2, strDate)
                    
                    strName = Trim(aItem(2)): strSample = Trim(aItem(3))
                    '判断是否无主标本
                      
                    '-------------------------------------------------------------------------------
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
                            " And a.仪器ID=[3] And a.标本序号=[4] And Nvl(a.标本类别,0)=[5]"
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
                        gstrSQL = "Select a.*,Decode(A.性别,Null,0,'男',1,'女',2,0) As 性别A,to_char(c.出生日期,'yyyy-mm-dd') As 出生日期 From 检验标本记录 a,病人医嘱记录 b,病人信息 c " & _
                            " Where a.医嘱id=b.id(+) And b.病人id=c.病人id(+)" & _
                            " And a.核收时间 Between [1] And [2]" & _
                            " And a.仪器ID=[3] And a.标本序号=[4] And Nvl(a.标本类别,0)=[5] and a.标本序号 = [6] "
                        Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "查询标本记录", CDate(strStartDate), _
                            CDate(strEndDate), lngDeviceID, strSampleID, iType, strSampleID)
                    End If
                    '-------------------------------------------------------------------------------
                    If rsTmp.EOF Then
                        '无主标本增加临时标本记录
                        strSex = -1
                        strBirth = ""
                        lngID = gobjDatabase.GetNextId("检验标本记录")
                        gstrSQL = "ZL_检验标本记录_INSERT(" & lngID & ",NULL,'" & _
                            strSampleID & "',NULL,NULL," & lngDeviceID & ",NULL," & _
                            "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),NULL," & _
                            "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & strSample & "'," & _
                            "Null,To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & strName & "','" & aItem(4) & "'," & lngExeDeptID & "," & iType & "," & intMicrobe & ")"
                        ExecuteProcedure "插入检验临时记录", gcnOracle
                    Else
                        strSex = Nvl(rsTmp("性别A"), 0)
                        strBirth = Nvl(rsTmp("出生日期"))
                        If intMicrobe = 0 Then
                            strSample = Nvl(rsTmp("标本类型"))
                        End If
                        lngID = rsTmp("ID")
                        blnAuditing = Not IsNull(rsTmp("审核人"))
                    End If
                        
                    If Not blnAuditing Then
                        '处理检验项目
                        strItemRecords = ""
                        str未知项 = ""
                        For j = 5 To UBound(aItem) Step 2
                            '根据通道号修改相应项目结果，未找到的则直接增加（根据通道号找不到项目的暂不处理）
                            '根据通道号找项目
                            lngItemID = GetItemID(aItem(j))
                            If lngItemID > 0 Then
                                strItemRecords = strItemRecords & "|" & lngItemID & "^" & aItem(j + 1)
                            Else
    
                                If str未知项 = "" Then str未知项 = "标本号    项目标识    项目值" & vbNewLine
                                str未知项 = str未知项 & strSampleID & vbTab & aItem(j) & vbTab & aItem(j + 1) & vbNewLine
    '                            gcnAccess.Execute strSql
                            End If
                        Next
                        If str未知项 <> "" Then Call WriteLog("mdlLISComm.ResultFromFile", LOG_未知项, 0, str未知项)
                        If Len(strItemRecords) > 0 Then
                            strItemRecords = Mid(strItemRecords, 2)
                            
                            gstrSQL = "ZL_检验普通结果_BATCHUPDATE(" & lngID & "," & _
                                lngDeviceID & ",'" & strSample & "'," & strSex & "," & _
                                IIf(strBirth = "", "Null", "To_Date('" & strBirth & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                                strItemRecords & "'," & intMicrobe & ")"
                            ExecuteProcedure "检验结果报告", gcnOracle
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
            End If
        Next
    End If
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
'    If gcnAccess.State <> adStateClosed Then gcnAccess.CommitTrans
    Exit Sub
DBError:
    If blnBeginTrans Then gcnOracle.RollbackTrans
    Call WriteLog("mdlLISComm.ResultFromFile", LOG_错误日志, Err.Number, Err.Description & vbCrLf & gstrSQL)
End Sub

Private Function GetItemID(ByVal strChannel As String) As Long
    Dim i As Integer
    For i = 0 To UBound(mItem, 2)
        If UCase(strChannel) = UCase(mItem(0, i)) Then Exit For
    Next
    If i > UBound(mItem, 2) Then
        GetItemID = -1
    Else
        GetItemID = CLng(mItem(1, i))
    End If
End Function

Public Function ValidateIP(ByVal strIP As String, Optional strErrInfo As String) As Boolean
    '检查IP地址的正确性。
    
    Dim varIP As Variant
    Dim IPError As Integer
    Dim IPd As Integer
    Dim i As Integer
    
    varIP = Split(strIP, ".")
    If UBound(varIP) <> 3 Then
        IPError = 0
    Else
        For i = 0 To 3
            If Not IsNumeric(varIP(i)) Then
                IPError = 1
                Exit For
            Else
                IPd = CInt(varIP(i))
                If IPd < 0 Or IPd > 255 Then
                    IPError = 2
                    Exit For
                Else
                    IPError = -1
                End If
            End If
        Next i
    End If
    
    ValidateIP = True
    Select Case IPError
        Case -1
            If strIP <> "0.0.0.0" Then
                ValidateIP = False
                strErrInfo = ""
            Else
                strErrInfo = "IP不能设为0.0.0.0。"
            End If
        Case 0
            strErrInfo = "IP格式不对，应为XXX.XXX.XXX.XXX。其中XXX为0-255的数字。"
        Case 1
            strErrInfo = "IP地址只能为0-255的数字。"
        Case 2
            strErrInfo = "IP地址的范围只能为0-255之间。"
    End Select
End Function

Public Function ValidatePort(ByVal strPort As String, Optional strErrInfo As String) As Boolean
    '检查端口号的正确性。
    ValidatePort = True
    If Not IsNumeric(Trim(strPort)) Then
        strErrInfo = "端口号只能为1-65535的数字。"
    Else
        If Val(Trim(strPort)) > 0 And Val(Trim(strPort)) <= 65535 Then
            ValidatePort = False
            strErrInfo = ""
        Else
            strErrInfo = "端口号的范围只能在1-65535之间。"
        End If
    End If
End Function

Private Sub WriteTxtLog(ByVal lng类型 As String, ByVal str项目 As String, ByVal str内容 As String)
    '以下变量用于记录调用接口的入参
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Dim blnClearData As Boolean
    
    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    'If Val(GetSetting("ZLSOFT", "zlLisLog", "Test", 0)) = 0 Then Exit Sub
    
    blnClearData = gblnClearData
    
    '错误日志(产生时间,错误类型,错误号,错误信息
    If str项目 <> "" Or str内容 <> "" Then
        
        If lng类型 = LOG_错误日志 Then
            '错误日志
            strFileName = App.Path & "\zlLis错误日志_" & Format(date, "yyyyMMdd") & ".LOG"
            If pLast错误日志 = str项目 & "|" & str内容 Then
                Exit Sub
            Else
                pLast错误日志 = str项目 & "|" & str内容
            End If
        ElseIf lng类型 = LOG_通讯日志 Then
            '通讯日志
            
            If blnClearData Then Exit Sub '勾了清空日志选项，则不写日志
            strFileName = App.Path & "\zlLis通讯日志_" & Format(date, "yyyyMMdd") & ".LOG"
            If pLast通讯日志 = str项目 & "|" & str内容 Then
                Exit Sub
            Else
                pLast通讯日志 = str项目 & "|" & str内容
            End If
        ElseIf lng类型 = LOG_未知项 Then
            '未知项
            If blnClearData Then Exit Sub '勾了清空日志选项，则不写日志
            strFileName = App.Path & "\zlLis未知项目_" & Format(date, "yyyyMMdd") & ".LOG"
        End If
        
        If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
        Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
        
        
        strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
        objStream.WriteLine ("时间:" & strDate & " 版本:" & App.major & "." & App.minor & "." & App.Revision)
        
        objStream.WriteLine (str项目)
        objStream.WriteLine (str内容)
        
        'objStream.WriteLine (String(50, "-"))
        objStream.Close
        Set objStream = Nothing
    End If
End Sub

Public Sub SaveImg(ByVal lngID As Long, ByVal strImg As String)
    '保存图形数据到数据库中
    
    Dim aGraphItem() As String
    Dim strImageVal As String
    Dim strImageType As String
    Dim strImageData As String
    Dim intLoop As Integer
    Dim IntCount As Integer
    Dim blnDeleImg As Boolean '保存后是否删除原来的图片
    Dim strPicPath As String, strSQL() As String
    Dim intLayOut As Integer '图片的显示方式
    Dim strBMPFile As String
    
    On Error GoTo ErrHandle
    aGraphItem = Split(strImg, "^")
        
    For intLoop = 0 To UBound(aGraphItem)
        strImageVal = Replace(aGraphItem(intLoop), vbCrLf, "")
        strImageType = Mid(strImageVal, 1, InStr(strImageVal, ";") - 1)
        strImageData = Mid(strImageVal, InStr(strImageVal, ";") + 1)
        
        If Mid(strImageData, 1, InStr(strImageData, ";") - 1) >= 100 And Mid(strImageData, 1, InStr(strImageData, ";") - 1) <= 227 Then
            '组织图片数据
            intLayOut = Mid(strImageData, 1, InStr(strImageData, ";") - 1)
            strPicPath = Mid(strImageData, InStr(strImageData, ";") + 1)
            
            If InStr(strPicPath, ";") > 0 Then
                strPicPath = Mid(strPicPath, InStr(strPicPath, ";") + 1)
                If Left(strPicPath, 2) = "1;" Then
                    blnDeleImg = True
                End If
            End If
            
            If Dir(strPicPath) <> "" Then
                If UCase(Right(strPicPath, 4)) = ".BMP" And intLayOut >= 100 And intLayOut <= 107 Then
                    strBMPFile = strPicPath
                ElseIf (UCase(Right(strPicPath, 4)) = ".JPG" Or UCase(Right(strPicPath, 4)) = ".GIF") And intLayOut >= 110 And intLayOut <= 127 Then
                    strBMPFile = strPicPath
                ElseIf intLayOut >= 200 And intLayOut <= 227 Then
                    strPicPath = UCase$(strPicPath)
                    strBMPFile = zlFileZip(strPicPath)
                Else
                    frmLISSrv.picTmp.Picture = LoadPicture(strPicPath)
                    If Dir(App.Path & "\zlLisIn.bmp") <> "" Then Kill App.Path & "\zlLisIn.bmp"
                    SavePicture frmLISSrv.picTmp.Picture, App.Path & "\zlLisIn.bmp"
                    strBMPFile = App.Path & "\zlLisIn.bmp"
                End If
                
                
                If zlLisBlobSql(lngID, strImageType, strBMPFile, intLayOut, strSQL) Then
                    WriteLog "执行 SaveImg", LOG_通讯日志, 0, "开始时间"
                    For IntCount = LBound(strSQL) To UBound(strSQL)
                        If strSQL(IntCount) <> "" Then
                            gstrSQL = strSQL(IntCount)
                            ExecuteProcedure Replace(strSQL(IntCount), "Call", ""), gcnOracle
                        End If
                    Next
                    WriteLog "执行 SaveImg", LOG_通讯日志, 0, "结束时间"
                End If
                If blnDeleImg Then
                    Kill strPicPath
                End If
                If intLayOut >= 200 And intLayOut <= 227 Then
                    Kill strBMPFile
                End If
            End If
        Else
            '图形数据
            If Len(strImageData) > 2000 Then
                '保存大于2000以上数据

                For IntCount = 1 To CInt(Len(strImageData) / 1000) + 1
                    If Len(strImageData) > 0 Then
                        
                        gstrSQL = "Zl_检验图像结果_Update(" & lngID & ",'" & strImageType & "','" & _
                                                Mid(strImageData, IntCount * 1000 - 999, 1000) & "'," & _
                                                "1," & IntCount & ")"
                        ExecuteProcedure "检验图像保存", gcnOracle
                    End If
                Next

            Else
                gstrSQL = "Zl_检验图像结果_Update(" & lngID & ",'" & strImageType & "','" & strImageData & "',0,1)"
                ExecuteProcedure "检验图像保存", gcnOracle
            End If
        End If
    Next

    Exit Sub
ErrHandle:
    Call WriteLog("SaveImg", LOG_错误日志, Err.Number, Err.Description)

End Sub


Private Function zlLisBlobSql(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String, ByVal layOut As Integer, ByRef arySql() As String) As Boolean
    '生成保存图片的SQL
    'Action 检验ID
    'KeyWord 标题
    'strFile 图片文件
    'arySql 生成的SQL存放在此数组中
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    Dim lngLBound As Long, lngUBound As Long    '传入数组的最小最大下标
    Err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUBound = UBound(arySql)
    If Err <> 0 Then lngLBound = 0: lngUBound = -1
    Err = 0: On Error GoTo 0
    
    lngFileNum = FreeFile
    WriteLog "生成BlobSQL", LOG_通讯日志, 0, "开始时间"
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    Err = 0: On Error GoTo errHand
    conChunkSize = 512
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    
    ReDim Preserve arySql(lngLBound To lngUBound + lngBlocks + 1)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        strText = Join(aryHex, "")
        If strText <> "" Then
            If lngCount = 0 Then strText = layOut & ";" & strText
            arySql(lngUBound + lngCount + 1) = "Zl_检验图像结果_Update(" & Action & ",'" & KeyWord & "','" & strText & "',1," & IIf(lngCount = 0, 1, 0) & ")"
        End If
    Next
    Close lngFileNum
    WriteLog "生成BlobSQL", LOG_通讯日志, 0, "结束时间"
    zlLisBlobSql = True
    Exit Function

errHand:
    Close lngFileNum
    zlLisBlobSql = False
End Function
Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1, Optional ByVal BeginDate As String) As String
    '-----------------------------------------------------------------------------------------
    '功能:获取特殊时间
    '参数:
    '-----------------------------------------------------------------------------------------
    Dim intDay As Integer
    Dim dateNow As Date
    
    If BeginDate = "" Then
        dateNow = gobjDatabase.Currentdate
    Else
        dateNow = BeginDate
    End If
    
    Select Case strMode
    Case "当  时"      '当时
        GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(dateNow, "YYYY-MM-DD")))
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 2, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 8 - intDay, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(dateNow, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(dateNow, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(dateNow, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "不重复"
        If bytFlag = 1 Then
            GetDateTime = "2000-01-01 00:00:00"
        Else
            GetDateTime = "3000-12-31 23:59:59"
        End If
    End Select
    
End Function

Public Function CreateSample(ByVal lngDeviceID As Long, ByVal strBarcode As String, _
    ByRef strSampleNO As String, ByVal dtSampleDate As Date, ByVal intType As Integer) As Boolean
    'inttype=0
    Dim strSQL As String, rsTmp As adodb.Recordset, rs As New adodb.Recordset
    Dim lngKey As Long, strItemRecords As String
    Dim lngDeptID As Long '当前仪器科室
    Dim rsItem As New adodb.Recordset
    Dim strItem As String                           '检验项目
    Dim str姓名 As String, str性别 As String, str年龄 As String
    On Error GoTo DBErr
    
    CreateSample = False
    
    '查找仪器科室
    strSQL = "Select 使用小组id From 检验仪器 Where ID = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "生成条码标本", lngDeviceID)
    lngDeptID = lngExeDeptID
    If Not rsTmp.EOF Then
        lngDeptID = Nvl(rsTmp("使用小组id"), lngExeDeptID)
    End If
    
    If Val(strSampleNO) <= 0 Then
        strSampleNO = Val(CalcNextCode(lngDeviceID, 0, intType))
    End If

    '查找符合条码的项目指标
'    strSql = "Select A.相关ID AS ID," & _
        "C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)') As 姓名,A.性别,A.年龄,F.No," & _
        "I.诊治项目ID As 项目ID,Decode(I.结果类型,3,Nvl(I.默认值,'-'),2,I.默认值,'') As 结果,'' As 标志," & _
        "Trim(REPLACE(REPLACE(' '||zlGetReference(I.诊治项目ID,A.标本部位,DECODE(A.性别,'男',1,'女',2,0),C.出生日期,Y.仪器ID,A.年龄),' .','0.'),'～.','～0.')) AS 结果参考," & _
        "NVL(A.紧急标志,0) AS 紧急,F.采样时间,F.采样人 " & _
        "FROM 病人医嘱记录 A," & _
        "病人信息 C,病人医嘱发送 F,检验报告项目 G,检验项目 I,检验仪器项目 Y " & _
        "WHERE A.诊疗类别 = 'C' " & _
        "AND A.病人ID=C.病人ID " & _
        "AND A.相关id IS NOT NULL " & _
        "AND A.医嘱状态=8 AND A.ID=F.医嘱id " & _
        "AND A.诊疗项目id=G.诊疗项目id AND G.细菌ID Is Null AND G.报告项目id=Y.项目id(+) " & _
        "AND G.报告项目ID=I.诊治项目ID " & _
        "AND (Y.仪器ID+0=[1] Or (Y.仪器ID Is Null And F.执行部门ID=[3])) " & _
        "And F.样本条码=[2] "
'        "AND F.执行状态=0 "
    
    strSQL = "Select ID, 姓名, 性别, 年龄, NO, 项目id, 结果, 标志, 结果参考, 紧急, 采样时间, 采样人, Rownum As 排列序号, 诊疗项目id," & vbNewLine & _
            "       编码,标本部位,开嘱科室ID,开嘱医生,标识号,当前床号,病人科室 " & vbNewLine & _
            "From (Select A.相关id As ID, C.姓名 || Decode(A.婴儿, 0, '', Null, '', '(婴儿)') As 姓名, A.性别, A.年龄, F.NO," & vbNewLine & _
            "              I.诊治项目id As 项目id, Decode(I.结果类型, 3, Nvl(I.默认值, '-'), 2, I.默认值, '') As 结果, '' As 标志," & vbNewLine & _
            "              Trim(Replace(Replace(' ' || Zlgetreference(I.诊治项目id, A.标本部位, Decode(A.性别, '男', 1, '女', 2, 0)," & vbNewLine & _
            "                                                          C.出生日期, Y.仪器id, A.年龄), ' .', '0.'), '～.', '～0.')) As 结果参考," & vbNewLine & _
            "              Nvl(A.紧急标志, 0) As 紧急, F.采样时间, F.采样人, G.排列序号, A.诊疗项目id, M.编码, " & vbNewLine & _
            "              a.标本部位,开嘱科室ID,开嘱医生,decode(a.病人来源,2,c.住院号,c.门诊号) as 标识号,c.当前床号,l.名称 as 病人科室 " & vbNewLine & _
            "       From 病人医嘱记录 A, 病人信息 C, 病人医嘱发送 F, 检验报告项目 G, 检验项目 I, 检验仪器项目 Y, 诊疗项目目录 M ,部门表 L " & vbNewLine & _
            "       Where A.诊疗类别 = 'C' And A.病人id = C.病人id And A.相关id Is Not Null And A.医嘱状态 = 8 And A.ID = F.医嘱id And" & vbNewLine & _
            "             A.诊疗项目id = G.诊疗项目id And G.细菌id Is Null And G.报告项目id = Y.项目id(+) And" & vbNewLine & _
            "             G.报告项目id = I.诊治项目id And A.诊疗项目id = M.ID(+) And a.病人科室ID = l.ID" & vbNewLine & _
            "             and (Y.仪器id + 0 = [1] Or (Y.仪器id Is Null And F.执行部门id = [3])) And nvl(F.执行状态,0) = 0  And F.样本条码 = [2]" & vbNewLine & _
            "       Order By M.编码, G.排列序号)"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "生成条码标本", lngDeviceID, strBarcode, lngDeptID)
    If rsTmp.EOF Then Exit Function
    
    gstrSQL = "Select B.病人id, B.主页id, B.序号, B.婴儿姓名, B.婴儿性别" & vbNewLine & _
                    "From 病人医嘱记录 A, 病人新生儿记录 B" & vbNewLine & _
                    "Where A.病人id = B.病人id And A.主页id = B.主页id And A.婴儿 = B.序号 And A.相关id = [1] And Rownum = 1"
    Set rs = gobjDatabase.OpenSQLRecord(gstrSQL, "CreateSample", CLng(rsTmp("ID")))
    If rs.EOF = False Then
        str姓名 = Nvl(rs("婴儿姓名"))
        str性别 = Nvl(rs("婴儿性别"))
        str年龄 = "婴儿"
    Else
        str姓名 = Nvl(rsTmp("姓名"))
        str性别 = Nvl(rsTmp("性别"))
        str年龄 = Nvl(rsTmp("年龄"))
    End If
    
    '读出检验项目
    gstrSQL = "select distinct 医嘱内容 from 病人医嘱记录 a , 病人医嘱发送 b, 检验报告项目 c , 检验仪器项目 d " & vbNewLine & _
              "  where a.id = b.医嘱ID and a.相关id is not null and a.诊疗项目ID = c.诊疗项目ID and " & vbNewLine & _
              "  c.报告项目ID = d.项目ID(+) and  (d.仪器id + 0 = [1] Or (d.仪器id Is Null And b.执行部门id = [3])) and b.样本条码 = [2] "
    Set rsItem = gobjDatabase.OpenSQLRecord(gstrSQL, "生成条码标本_1", lngDeviceID, strBarcode, lngDeptID)
    Do Until rsItem.EOF
        strItem = strItem & " " & Nvl(rsItem("医嘱内容"))
        rsItem.MoveNext
    Loop
    strItem = Trim(strItem) & "(" & Nvl(rsTmp("标本部位")) & ")"
        
    '产生标本记录
    lngKey = gobjDatabase.GetNextId("检验标本记录")
    gstrSQL = "ZL_检验标本记录_标本核收(" & lngKey & "," & _
        rsTmp("ID") & ",'" & _
        strSampleNO & "'," & _
        IIf(IsNull(rsTmp("采样时间")), "Null", "TO_DATE('" & Format(rsTmp("采样时间"), "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')") & "," & _
        IIf(IsNull(rsTmp("采样人")), "Null", "'" & rsTmp("采样人") & "'") & "," & _
        lngDeviceID & "," & _
        "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),Null," & _
        "1,'" & _
        gobjDatabase.GetUserInfo.Fields("姓名").Value & "'," & _
        "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),0,0,0," & _
        intType & ",NULL,'" & _
        str姓名 & "','" & str性别 & "','" & str年龄 & "','" & Nvl(rsTmp("No")) & "','" & _
        Nvl(rsTmp("标本部位")) & "'," & Nvl(rsTmp("开嘱科室ID")) & ",'" & Nvl(rsTmp("开嘱医生")) & "'," & _
        Nvl(rsTmp("标识号")) & ",'" & Nvl(rsTmp("当前床号")) & "','" & Nvl(rsTmp("病人科室")) & "','" & _
        strItem & "')"
    ExecuteProcedure "生成条码标本", gcnOracle
                                                                
    '填写指标
    strItemRecords = ""
    Do While Not rsTmp.EOF
        strItemRecords = strItemRecords & "|" & rsTmp("ID") & "^" & rsTmp("项目ID") & "^" & _
            Nvl(rsTmp("结果")) & "^" & Nvl(rsTmp("标志"), 0) & "^" & Nvl(rsTmp("结果参考")) & "^" & _
            Nvl(rsTmp("诊疗项目ID")) & "^" & Nvl(rsTmp("排列序号"))
            
        rsTmp.MoveNext
    Loop
    
    If Len(strItemRecords) > 0 Then
        strItemRecords = Mid(strItemRecords, 2)
            
        gstrSQL = "Zl_检验普通结果_Write(" & lngKey & "," & _
            lngDeviceID & ",'" & strItemRecords & "',0,0)"
        ExecuteProcedure "生成条码标本", gcnOracle
    End If
    Exit Function
DBErr:
    Call WriteLog("clsLISComm.CreateSample", LOG_错误日志, Err.Number, Err.Description)
End Function

Private Function CalcNextCode(ByVal lngKey As Long, ByVal intRow As Integer, ByVal iType As Integer) As String
    '--------------------------------------------------------------------------------------------------------
    '功能:计算指定仪器在当天内的下一个缺省标本号
    '参数:lngKey                检验仪器ID
    '     iType                 标本类别：0=普通、1=急诊
    '返回:缺省标本号码
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New adodb.Recordset
    Dim strToday As String
    Dim strTmp As String
    Dim lng次数 As Long
    Dim strLabNo As String, strLabQCNo As String '检验标本、质控标本
    Dim mstrSQL As String, mlngLoop As Long
    Dim mlngDefaultItemID As Long
    
    '时间,仪器,标本号
    On Error GoTo errHand
    mlngDefaultItemID = 0
    strToday = Format(gobjDatabase.Currentdate, "YYYY-MM-DD")
    
    On Error GoTo point1
    
    mstrSQL = "SELECT NVL(MAX(TO_NUMBER(标本序号)),0) AS 最大序号 FROM 检验标本记录 a,检验申请项目 b " & _
                "WHERE 核收时间 BETWEEN [2] and [3] And a.id = b.标本id(+) And nvl(a.是否质控品,0) = 0 " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL " & _
                        IIf(mlngDefaultItemID > 0, " And b.诊疗项目id = [4] ", ""), "AND 仪器id= [1] ") & " And 医嘱ID Is Not Null" & _
                    IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1")
    Set rs = gobjDatabase.OpenSQLRecord(mstrSQL, "计算", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                           CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabNo = gobjCommFun.Nvl(rs("最大序号"))
    
    On Error GoTo errHand
    GoTo point2
    
point1:
    On Error GoTo errHand
    
    mstrSQL = "SELECT NVL(MAX(标本序号),'') AS 最大序号 FROM 检验标本记录 a,检验申请项目 b " & _
                "WHERE 核收时间 BETWEEN [2] and [3] And a.id = b.标本id(+) And nvl(a.是否质控品,0) = 0 " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL " & _
                    IIf(mlngDefaultItemID > 0, " And b.诊疗项目id = [4] ", ""), "AND 仪器id= [1] ") & " And 医嘱ID Is Not Null" & _
                    IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1")
    Set rs = gobjDatabase.OpenSQLRecord(mstrSQL, "计算", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                            CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabNo = gobjCommFun.Nvl(rs("最大序号"))
    
point2:
    On Error GoTo point3
    
    mstrSQL = "SELECT NVL(MAX(TO_NUMBER(标本序号)),0) AS 最大序号 FROM 检验标本记录 a,检验申请项目 b " & _
                "WHERE 核收时间 BETWEEN [2] and [3] And a.id = b.标本ID(+) And nvl(a.是否质控品,0) = 0 " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL " & _
                    IIf(mlngDefaultItemID > 0, " And b.诊疗项目id = [4] ", ""), "AND 仪器id= [1] ") & _
                    IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1")
    Set rs = gobjDatabase.OpenSQLRecord(mstrSQL, "计算", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                            CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabQCNo = gobjCommFun.Nvl(rs("最大序号"))
    
    On Error GoTo errHand
    GoTo point4
    
point3:
    On Error GoTo errHand
    
    mstrSQL = "SELECT NVL(MAX(标本序号),'') AS 最大序号 FROM 检验标本记录 a,检验申请项目 b" & _
                " WHERE 核收时间 BETWEEN [2] and [3] And a.id = b.标本ID(+) And nvl(a.是否质控品,0) = 0 " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL " & _
                    IIf(mlngDefaultItemID > 0, " And b.诊疗项目id = [4] ", ""), "AND 仪器id=[1] ") & _
                    IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1")
    Set rs = gobjDatabase.OpenSQLRecord(mstrSQL, "计算", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                            CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabQCNo = gobjCommFun.Nvl(rs("最大序号"))
    
point4:
    If strLabNo >= strLabQCNo Then
        CalcNextCode = strLabNo
    Else
        CalcNextCode = strLabQCNo
    End If
'    If Val(strLabQCNo) > Val(strLabNo) + 100 Then CalcNextCode = strLabNo

'    For mlngLoop = 1 To vsf2.Rows - 1
'        If mlngLoop <> intRow Then
'            If Val(vsf2.RowData(mlngLoop)) = lngKey Then
'                If Val(CalcNextCode) < Val(vsf2.TextMatrix(mlngLoop, 2)) Then
'                    CalcNextCode = Val(vsf2.TextMatrix(mlngLoop, 2))
'                End If
'            End If
'        End If
'    Next
'
    If Val(CalcNextCode) <= 0 Then
        CalcNextCode = "1"
        Exit Function
    End If
'
    CalcNextCode = Val(CalcNextCode) + 1
    Exit Function
    
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function
'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String, Optional ByVal strUnZipFile As String) As String
    Dim strZipPath As String
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If gobjFSO.FileExists(strUnZipFile) Then gobjFSO.DeleteFile strUnZipFile
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strUnZipFile) <> "" Then
        zlFileUnzip = strUnZipFile
    Else
        zlFileUnzip = ""
    End If
End Function
'################################################################################################################
'## 功能：  将文件压缩为新文件放到相同目录中
'## 参数：  strFile     :原始文件
'## 返回：  压缩文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLLIS" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
End Function
