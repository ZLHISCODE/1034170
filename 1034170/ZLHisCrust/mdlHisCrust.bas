Attribute VB_Name = "mdlHisCrust"
Option Explicit
'全局变量对象
Public gstrSetupPath        As String                   '程序的安装路径
Public garrKillProcess      As Variant                  '升级中杀掉的进程名称
Public gstrPreTempPath      As String                   '系统目录System32目录
Public gstrSystemPath       As String                   '系统目录System32目录
Public gstrTempPath         As String                   '临时存放目录
Public grsFileUpgrade       As ADODB.Recordset          '升级文件清单
Public gcnOracle            As ADODB.Connection
Public gstrComputerName     As String                   '电脑名称
Public gstrComputerIp       As String                   '本机的IP地址

Public gobjFSO              As New FileSystemObject     '文件操作对象
Public gobjTrace            As New clsTrace             '日志跟踪对象
Public gcllSetPath          As New Collection           '所有安装路径
Public gclsRegCom           As New clsRegCom            '部件注册对象
Public grsErrRec            As ADODB.Recordset          '错误记录
Public gclsConnect          As clsConnect               '文件收集的连接
Public gobj7zZip            As New cls7zZip             '7z压缩类

Public glngNoteLength       As Long                     '说明字段长度
Public glngFileBatch        As Long                     '升级文件批次
Private mblnWriteRunErrLog  As Boolean                  '是否书写运行错误日志，由数据库参数控制
Public gblnReCheckComs      As Boolean                  '是否重新检查安装部件
Public gintWaite            As Integer                  '等待升级的时间。0-立即升级，<>0等待N分钟后开始升级，一般为1
Public gblnIs64Bits         As Boolean                  '是否是64位系统
Public gblnHaveVersion      As Boolean                  '是否存在文件版本号字段
Public gblnSameFTP          As Boolean                  '是否使用简易FTP工具
'命令行解析内容
Public gstrCommand          As String                   '自动升级的命令行
Public gstrConnectString    As String                   '连接字符串
Public gotCurType           As OperateType              '本次执行的操作类型
Public gstrHisInput         As String                   'ZLHIS输入的用户名密码,格式为USER=ZLHIS PASS=HIS SERVER=TXYY(界面输入的密码)
Public gstrHisCommand       As String                   '原始的ZLHIS命令
Public gstrAppEXE           As String                   '调用本外壳程序的文件
Public gintCallTimes        As Integer                  '调用次数
Public gblnAutoLogin        As Boolean                  '是否自动登录
Public gstrTerminal         As String                   '当前机器名
Public gstrAudsid           As String                   '当前audsid
Public gblnSilence          As Boolean

Public gobjRegistry         As New clsRegistry


Private Sub Main()
    Dim objHisCrustCom  As Object
    Dim objHisCrust     As clsHisCrust
    Dim strReturn       As String
    Dim blnOLD          As Boolean
    
    On Error GoTo ErrH
    gblnAutoLogin = True
    gblnIs64Bits = Is64bit
    gstrSetupPath = GetSetupPath
    Call gobjTrace.OpenTace("ZLHISCRUST", gstrSetupPath)
    gobjTrace.WriteSection "客户端自动升级"
    gobjTrace.WriteSection "环境初始化", SL_LevelTwo
    
    gobjTrace.LogFile = gobjTrace.LogFile
    gstrCommand = GetCommand()
    If gstrCommand = "" Then GoTo ReCall
    gstrTerminal = InitTerminal(gstrAudsid)
    If Not GetBaseInfo Then GoTo ReCall
    '检查任务
    If Not CheckJobs Then
        GoTo ReCall
    ElseIf gclsConnect Is Nothing Then                   '没有任务，自动退出，登录ZLHIS
        GoTo AutoLogin
    End If

    Call EnablePrivilege(GetCurrentProcess(), SE_DEBUG_NAME)
    If Not SetOperateProcess(gotCurType, OS_InProcessing, SumErrMsg) Then GoTo ReCall
    '安装路径修复
    If Not CheckAndAdjustFolder() Then
        Call RecordErrMsg(MT_MsgFoot, "消息尾", "结果:升级失败 时间:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
        Call SetOperateProcess(gotCurType, OS_Failure, SumErrMsg)          '标识升级结束
        GoTo ReCall
    End If
    If gotCurType <> OT_CheckFile Then
        '剩余空间检查
        If Not CheckFreeSpace Then
            Call RecordErrMsg(MT_MsgFoot, "消息尾", "结果:升级失败 时间:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
            Call SetOperateProcess(gotCurType, OS_Failure, SumErrMsg)          '标识升级结束
            GoTo ReCall
        End If
        '升级基础部件
        If Not UpgradeBase() Then
            Call RecordErrMsg(MT_MsgFoot, "消息尾", "结果:升级失败 时间:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
            Call SetOperateProcess(gotCurType, OS_Failure, SumErrMsg)          '标识升级结束
            GoTo ReCall                      '强制退出进程
        Else
            Set objHisCrustCom = GetCrustObject
            If Not objHisCrustCom Is Nothing Then
                Set objHisCrust = New clsHisCrust
                Call gobjTrace.CloseLog
                Call objHisCrustCom.InitHisCrustCom(objHisCrust, IIf(IsDesinMode, "C:\APPSOFT\ZLHISCRUST.EXE", App.Path), CStr(Command()), gobjTrace.LogFile)
                Call objHisCrustCom.UpdateBaseFile
                Call objHisCrustCom.UpdateFile
                Exit Sub
            End If
        End If
    End If
    '获取升级文件，失败则强制退出
    If Not GetUpgradeFileList Then
        Call RecordErrMsg(MT_MsgFoot, "消息尾", "结果:升级失败 时间:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
        Call SetOperateProcess(gotCurType, IIf(gotCurType = OT_CheckFile, OS_Failure, OS_NotInProcessing), SumErrMsg)          '标识升级结束
        GoTo ReCall
    End If
    If grsFileUpgrade.RecordCount = 0 Then
        Call RecordErrMsg(MT_InitEnv, "文件清单获取", "没有可升级的文件，系统自动退出。")
        Call RecordErrMsg(MT_MsgFoot, "消息尾", "结果:升级完成 时间:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
        Call SetOperateProcess(gotCurType, OS_Completed, SumErrMsg, glngFileBatch)          '标识升级结束
        GoTo ReCall
    End If
    Call GetKILLProcess
    If gotCurType = OT_PreUpgrade Or gotCurType = OT_CheckFile Then
        frmHisCrust.Hide
    Else
        frmHisCrust.Show
    End If
    Exit Sub
ErrH:
    MsgBox Err.Description, vbInformation, App.Title
    Err.Clear
ReCall:
    gobjTrace.LogFile = gobjTrace.LogFile
    Call CallHISEXE
    End
AutoLogin:
    gobjTrace.LogFile = gobjTrace.LogFile
    Call CallHISEXE(True)
    End
End Sub

Private Function GetCrustObject() As Object
    On Error Resume Next
    Set GetCrustObject = CreateObject("zlHisCrustCom.clsHisCrsutCom")
    If Err.Number <> 0 Then
        Err.Clear
        If gobjFSO.FileExists(gstrSetupPath & "\ZLHISCRUSTCOM.DLL") Then
            Call gclsRegCom.RegCom(gstrSetupPath & "\ZLHISCRUSTCOM.DLL")
            Set GetCrustObject = CreateObject("zlHisCrustCom.clsHisCrsutCom")
            If Err.Number <> 0 Then
                Err.Clear
            End If
        End If
    End If
    '用来调试
    If Val(GetSetting("ZLSOFT", "公共模块\自动升级", "不启用COM", "0")) = 1 Then
        Set GetCrustObject = Nothing
    End If
End Function

Private Function GetSetupPath() As String
'功能：获取程序的安装路径
    If IsDesinMode Then
        GetSetupPath = "C:\APPSOFT"
    Else
        '可能以前放在Apply，但是由于可能被杀毒软件放入隔离区再次处理会失败
        '因此增加ZLuptmp升级目录，子目录为年份月份日期+时间，防止报杀。
        '如2016-12-12 12:12目录为APPSost\ZLUpTmp\1612121212
        '以前ZLHISCrust.EXE放在APPLY,新方式，放在APPSOFT\ZLUPTMP,解压同时也放在此处APPSOFT\ZLUPTMP
        If InStr(UCase(App.Path), "\ZLUPTMP") > 0 Then
            GetSetupPath = gobjFSO.GetParentFolderName(gobjFSO.GetParentFolderName(App.Path))
        ElseIf InStr(UCase(App.Path), "\APPLY") > 0 Then
            GetSetupPath = gobjFSO.GetParentFolderName(App.Path)
        Else
            GetSetupPath = App.Path
        End If
    End If
End Function

Private Function GetCommand() As String
    Dim strCommand      As String, strServer        As String
    Dim objText         As TextStream
    Dim strErrInfo      As String
    
    On Error GoTo ErrH
    gobjTrace.WriteSection "获取连接", SL_LevelThree
    strCommand = Command
    gobjTrace.WriteInfo "GetCommand", "原始启动命令行", Cipher(strCommand)
    'ZLRunAS.exe调用没有命令行,通过本地文件传递原始命令行
    If strCommand = "" Then
        If gobjFSO.FileExists(gstrSetupPath & "\ZLRUNAS.ini") Then
            'ZLRunAS.exe调用没有命令行
            Set objText = gobjFSO.OpenTextFile(gstrSetupPath & "\ZLRUNAS.ini", ForReading)
            If Not objText.AtEndOfLine Then
                strCommand = objText.ReadLine
            End If
            objText.Close
            Set objText = Nothing
            Call gobjFSO.DeleteFile(gstrSetupPath & "\ZLRUNAS.ini", True)
            gobjTrace.WriteInfo "GetCommand", "ZLRUNAS启动命令行", strCommand
            strCommand = DeCipher(strCommand)
        End If
    End If
    '通过配置文件生成加密串
    If strCommand = "" Then
        If gobjFSO.FileExists(gstrSetupPath & "\ZLHISCRUST.ini") Then
            Set objText = gobjFSO.OpenTextFile(gstrSetupPath & "\ZLHISCRUST.ini", ForReading)
            If Not objText.AtEndOfLine Then
                strCommand = Trim(objText.ReadLine)
            End If
            objText.Close
            Set objText = Nothing
            Call gobjFSO.DeleteFile(gstrSetupPath & "\ZLHISCRUST.ini", True)
            If strCommand Like "ZLUPDATE:*" Then
            Else
                strCommand = "ZLUPDATE:" & Cipher(strCommand)
            End If
            gobjTrace.WriteInfo "GetCommand", "配置启动命令行", strCommand
        End If
    End If
    '没有命令行
    If strCommand = "" Then
        If IsDesinMode Then
'            strCommand = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=TESTBASE"";Persist Security Info=True;User ID=ZLHIS;Password=HIS;Data Provider=MSDASQL||0||OfficialUpgrade||||USER=ZLHIS PASS=aqa||CMDCHECK:1,171,174,191,193,214"
            strCommand = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=TESTBASE_LF"";Persist Security Info=True;User ID=ZLHIS;Password=HIS;Data Provider=MSDASQL||0||OfficialUpgrade||||USER=ZLHIS PASS=aqa||"
            gobjTrace.WriteInfo "GetCommand", "源码启动命令行", strCommand
        End If
    End If
    If strCommand = "" Then
        strServer = GetSetting(appName:="ZLSOFT", Section:="注册信息\登陆信息", Key:="SERVER", Default:="")
        If MsgBox("是否需要强制升级？", vbInformation + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
            Exit Function
        End If
        If strServer = "" Then strServer = InputBox("请输入连接的服务器", "提示")
        If strServer = "" Then Exit Function
        '使用ZLTOOLS/ZLTOOLS登录
        strCommand = "ZLUPDATE:" & Cipher("USER=ZLTOOLS PASS=ZLTOOLS SERVER=" & strServer & " MODE=0")
        gobjTrace.WriteInfo "GetCommand", "强制启动(1)命令行", strCommand
        Set gcnOracle = GetConnByCommand(strCommand)
        '使用ZLTOOLS/ZLSOFT登录
        If gcnOracle Is Nothing Then
            strCommand = "ZLUPDATE:" & Cipher("USER=ZLTOOLS PASS=ZLSOFT SERVER=" & strServer & " MODE=0")
            gobjTrace.WriteInfo "GetCommand", "强制启动(2)命令行", strCommand
            Set gcnOracle = GetConnByCommand(strCommand)
        End If
        '用户输入密码登录
        If gcnOracle Is Nothing Then
            strCommand = InputBox("请输入ZLTOOLS的密码", "提示")
            If strCommand = "" Then Exit Function
            strCommand = "ZLUPDATE:" & Cipher("USER=ZLTOOLS PASS=" & strCommand & " SERVER=" & strServer & " MODE=0")
            gobjTrace.WriteInfo "GetCommand", "强制启动(3)命令行", strCommand
            Set gcnOracle = GetConnByCommand(strCommand, True)
        End If
    Else
        gobjTrace.WriteInfo "GetCommand", "常规启动命令行", Cipher(strCommand)
        Set gcnOracle = GetConnByCommand(strCommand, True)
    End If
    If gcnOracle Is Nothing Then Exit Function
    GetCommand = strCommand
    Exit Function
ErrH:
    strErrInfo = Err.Description
    gobjTrace.WriteInfo "GetCommand", "获取命令行失败", strErrInfo
    MsgBox "获取命令行信息发生致命错误，请联系管理员！信息：" & vbNewLine & strErrInfo, vbInformation, App.Title
    Err.Clear
End Function

Private Function GetConnByCommand(ByVal strCommand As String, Optional ByVal blnMsg As Boolean) As ADODB.Connection
'功能：通过命令行获取连接
'       strCommand=命令行
'       blnMsg=是否提示
'返回：创建的连接
    Dim strUser     As String, strPwd       As String, strServer    As String, intMode      As Integer
    Dim strTmp      As String, arrCommand   As Variant, i           As Integer
    Dim cnTmp       As ADODB.Connection
    Dim strCur      As String, lngWait      As Long
    
    On Error GoTo ErrH
    gstrHisInput = "": gstrHisCommand = "": gstrAppEXE = "": gintCallTimes = 0: gstrConnectString = "": gintWaite = 0
    If strCommand Like "ZLUPDATEEX:*" Then
        gobjTrace.WriteInfo "GetConnByCommand", "启动类型", "二次非常规启动"
        strCommand = DeCipher(Mid(strCommand, Len("ZLUPDATEEX:*")))
        gintCallTimes = 1
    End If
    
    '使用ZLHIS内置公用账户升级
    If strCommand Like "ZLUPDATE:*" Then
        arrCommand = Split(DeCipher(Mid(strCommand, Len("ZLUPDATE:*"))), " ")
        For i = LBound(arrCommand) To UBound(arrCommand)
            If arrCommand(i) Like "USER=*" Then
                strUser = Mid(arrCommand(i), Len("USER=*"))
            ElseIf arrCommand(i) Like "PASS=*" Then
                strPwd = Mid(arrCommand(i), Len("PASS=*"))
            ElseIf arrCommand(i) Like "SERVER=*" Then
                strServer = Mid(arrCommand(i), Len("SERVER=*"))
            ElseIf arrCommand(i) Like "MODE=*" Then
                intMode = Val(Mid(arrCommand(i), Len("MODE=*")))
            End If
        Next
        gblnAutoLogin = False
        If strServer = "" Then strServer = GetSetting(appName:="ZLSOFT", Section:="注册信息\登陆信息", Key:="SERVER", Default:="")
        If strUser = "" Or strPwd = "" Or strServer = "" Then
            gobjTrace.WriteInfo "GetConnByCommand", "启动失败", "命令行信息不完整，请联系管理!"
            If blnMsg Then
                MsgBox "命令行信息不完整，请联系管理员！", vbInformation, App.Title
            End If
            Exit Function
        End If
        gstrConnectString = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=" & strServer & """;Persist Security Info=True;User ID=" & strUser & ";Password=" & strPwd & ";Data Provider=MSDASQL"
        '升级类型
        gotCurType = Decode(intMode, 0, OT_Repair, 1, OT_OfficialUpgrade, 2, OT_PreUpgrade, 3, OT_CheckFile, OT_OfficialUpgrade)
    Else
        If strCommand Like "ZLUPDATENEW:*" Then
            gobjTrace.WriteInfo "GetConnByCommand", "启动类型", "二次常规启动"
            strCommand = DeCipher(Mid(strCommand, Len("ZLUPDATENEW:*")))
            gintCallTimes = 1
        End If
        arrCommand = Split(strCommand, "||")
        '校验方式进行命令解析，增加命令行解析的准确性
        If arrCommand(UBound(arrCommand)) Like "CMDCHECK:" Then
            strTmp = Mid(arrCommand(UBound(arrCommand)), 10)
            arrCommand = Split(strTmp, ",")
            strTmp = Mid(strCommand, 1, Len(strCommand) - Len(arrCommand(UBound(arrCommand))) - 2)
            For i = UBound(arrCommand) To LBound(arrCommand) Step -1
                If i = 5 Then
                    strCur = Mid(strTmp, Val(arrCommand(i)) + 2)
                    If strCur Like "W:*" Then '由于以前老方式的测试代码用For+Sleep函数实现等待，该方法存在程序崩溃问题，因此增加前缀W:
                        gintWaite = Val(Mid(strCur, 3))
                    End If
                ElseIf i = 4 Then               'HIS输入的用户名与密码
                    gstrHisInput = Mid(strTmp, Val(arrCommand(i)) + 2)
                ElseIf i = 3 Then
                    gstrHisCommand = Mid(strTmp, Val(arrCommand(i)) + 2)
                ElseIf i = 2 Then
                    gstrAppEXE = Mid(strTmp, Val(arrCommand(i)) + 2)
                ElseIf i = 1 Then
                    If gintCallTimes = 0 Then gintCallTimes = Val(Mid(strTmp, Val(arrCommand(i)) + 2))
                ElseIf i = 0 Then
                    gstrConnectString = strTmp
                End If
                strTmp = Mid(strTmp, 1, Val(arrCommand(i)) - 1)
            Next
        Else
            gstrConnectString = arrCommand(0)
            If gintCallTimes = 0 Then gintCallTimes = Val(arrCommand(1))
            gstrAppEXE = arrCommand(2)          '"PreUpgrade","OfficialUpgrade","zlActMain.exe"
            If UBound(arrCommand) >= 3 Then
                gstrHisCommand = arrCommand(3)
                If UBound(arrCommand) >= 4 Then
                    gstrHisInput = arrCommand(4)
                    If UBound(arrCommand) >= 5 Then
                        If arrCommand(5) Like "W:*" Then '由于以前老方式的测试代码用用For+Sleep函数实现等待，该方法存在程序崩溃问题，因此增加前缀W:
                            gintWaite = Val(Mid(arrCommand(5), 3))
                        End If
                    End If
                End If
            End If
        End If
        gotCurType = Decode(gstrAppEXE, "Repair", OT_Repair, "PreUpgrade", OT_PreUpgrade, "OfficialUpgrade", OT_OfficialUpgrade, "CheckUpgrade", OT_CheckFile, OT_OfficialUpgrade)
    End If
    gblnSilence = gotCurType = OT_CheckFile Or gotCurType = OT_PreUpgrade
    If gintWaite > 0 And gintCallTimes = 0 Then '只有第一次调用才沉睡
        lngWait = gintWaite * 60000
        Call Sleep(lngWait)
    End If
    Err.Clear: On Error Resume Next
    Set cnTmp = New ADODB.Connection
    cnTmp.CursorLocation = adUseClient
    cnTmp.ConnectionString = gstrConnectString
    cnTmp.Open
    If Err.Number <> 0 Then
        gobjTrace.WriteInfo "GetConnByCommand", "启动失败", Err.Description
        If blnMsg And Not gblnSilence Then
            MsgBox "打开数据库连接失败，请联系管理员！信息：" & vbNewLine & Err.Description, vbInformation, App.Title
        End If
        Err.Clear
        Exit Function
    End If
    gobjTrace.WriteInfo "GetConnByCommand", "任务", Decode(gotCurType, OT_Repair, "修复", OT_PreUpgrade, "预升级", OT_OfficialUpgrade, "正式升级", OT_CheckFile, "收集"), _
                "主调用程序", gstrAppEXE, "主调用程序命令", Cipher(gstrHisCommand), "主调程序输入", Cipher(gstrHisInput), "自我调用次数", gintCallTimes
    Set GetConnByCommand = cnTmp
    Exit Function
ErrH:
    gobjTrace.WriteInfo "GetConnByCommand", "启动获取连接失败", Err.Description
    If Not gblnSilence Then MsgBox "启动获取连接失败，请联系管理员！信息：" & vbNewLine & Err.Description, vbInformation, App.Title
    Err.Clear
End Function

Public Sub CallHISEXE(Optional blnAutoLogin As Boolean)
    '调用HIS
    Dim mError              As String
    Dim strFile             As String
    Dim strCommand          As String
    Dim strUserName         As String, strPass      As String, strServer As String
    Dim lngRet              As Long
    
    '如果是ZLBH融合启动，则不再回调
    If UCase(gstrAppEXE) = "ZLACTMAIN.EXE" Or gotCurType = OT_PreUpgrade Or gotCurType = OT_CheckFile Then Exit Sub
    '确定文件是否存在
    '1、不再处理 "ZLHIS90.exe"
    '2、预升级也会自动调用导航台程序
    If gstrAppEXE <> "" Then
        strFile = gstrSetupPath & "\" & gstrAppEXE
        If Not gobjFSO.FileExists(strFile) Then
            If UCase(gstrAppEXE) <> "ZLHIS+.EXE" Then
                strFile = gstrSetupPath & "\ZLHIS+.exe"
            End If
        End If
    Else
        strFile = gstrSetupPath & "\ZLHIS+.exe"
    End If
    gobjTrace.WriteInfo "CallHISEXE", "主调程序路径", strFile
    On Error Resume Next
    If blnAutoLogin And gblnAutoLogin Then
        '公共账户升级不自动登录
        If gstrHisCommand = "" And gstrHisInput = "" And Not (gstrCommand Like "ZLUPDATE:*" Or gstrCommand Like "ZLUPDATEEX:*") Then
            If GetConnectionInfo(gstrConnectString, strServer, strUserName, strPass) Then
                strCommand = strUserName & "/" & strPass & "@" & strServer
            End If
        ElseIf gstrHisCommand <> "" Then
            strCommand = gstrHisCommand
            If gstrHisCommand Like "USER=*" Then
                strCommand = gstrHisCommand & " ZLHISCRUSTCALL=1"
            End If
        Else
            strCommand = gstrHisInput & IIf(gstrHisInput <> "", " ZLHISCRUSTCALL=1", "")
        End If
        gobjTrace.WriteInfo "CallHISEXE", "命令行", Cipher(strCommand)
        strCommand = strFile & " " & strCommand
    Else
        strCommand = strFile
    End If
    lngRet = Shell(strCommand, vbNormalFocus)
    Call Sleep(100)
End Sub

Public Function GetConnectionInfo(ByVal strConect As String, ByRef strServerName As String, ByRef strUserName As String, ByRef strPassword As String) As Boolean
'功能： 分析MSODBC连接对象中的ORACLE串中的 服务器，用户名，密码
'返回： 成功失败，返回True；失败，返回False

    Dim i As Integer
    Dim strTemp As String
    If strConect = "" Then Exit Function
            
    strServerName = ""
    strUserName = ""
    strPassword = ""
    strConect = Replace(strConect, """", "")
    
    If InStr(strConect, "ODBC") > 0 Then
        'Provider=MSDataShape.1;Extended Properties="Driver={Microsoft ODBC for Oracle};Server=DYYY";Persist Security Info=True;User ID=zlhis;Password=his;Data Provider=MSDASQL"
        'Provider=MSDataShape.1;Persist Security Info=False;User ID=ZLHIS;Data Provider=MSDASQL;
        '获取 strServerName(Security为false时，无法获得)
        i = InStrRev(strConect, "Server=", -1)
        If i > 0 Then
            strTemp = Right(strConect, Len(strConect) - i - 6)
            i = InStr(1, strTemp, ";")
            If i > 0 Then
                strServerName = Left(strTemp, i - 1)
            End If
        End If
    Else
        'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source="DYYY";Extended Properties="PLSQLRSet=1"
        'Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=ZLHIS;Data Source="DYYY"
        i = InStrRev(strConect, "Data Source=", -1)
        If i > 0 Then
            strTemp = Right(strConect, Len(strConect) - i - 11)
            i = InStr(1, strTemp, ";")
            If i > 0 Then
                strServerName = Left(strTemp, i - 1)
            Else    'Security为false时，没有;号
                strServerName = strTemp
            End If
        End If
    End If
    
    '获取 strUserName
    i = InStrRev(strConect, "User ID=", -1)
    If i > 0 Then
        strTemp = Right(strConect, Len(strConect) - i - 7)
        i = InStr(1, strTemp, ";")
        If i > 0 Then
            strUserName = Left(strTemp, i - 1)
        End If
    End If
    
    '获取 strPassword
    i = InStrRev(strConect, "Password=", -1)
    If i > 0 Then
        strTemp = Right(strConect, Len(strConect) - i - 8)
        i = InStr(1, strTemp, ";")
        If i > 0 Then
            strPassword = Left(strTemp, i - 1)
        End If
    End If
    GetConnectionInfo = strPassword <> "" And strUserName <> "" And strServerName <> ""
End Function

Private Function GetBaseInfo() As Boolean
    Dim strErrInfo      As String
    
    On Error GoTo ErrH
    gstrComputerName = ComputerName
    gstrComputerIp = IP
    gstrSystemPath = gobjFSO.GetSpecialFolder(SystemFolder)
    If gblnIs64Bits Then '64系统下32位程序应该放在C:\windows\SysWOW64
        gstrSystemPath = gobjFSO.GetParentFolderName(gstrSystemPath) & "\SysWOW64"
    End If
    gblnReCheckComs = False
    gstrTempPath = gstrSetupPath & "\ZLUPTMP"
    If Not gobjFSO.FolderExists(gstrTempPath) Then
        Call gobjFSO.CreateFolder(gstrTempPath)
    End If
    gstrPreTempPath = gstrTempPath & "\ZLPRETMP"
    If Not gobjFSO.FolderExists(gstrPreTempPath) Then
        Call gobjFSO.CreateFolder(gstrPreTempPath)
    End If
    '临时目录加入动态目录
    gstrTempPath = gstrTempPath & "\" & Format(Now, "YYMMDDHHmmss")
    If Not gobjFSO.FolderExists(gstrTempPath) Then
        Call gobjFSO.CreateFolder(gstrTempPath)
    End If
    mblnWriteRunErrLog = IsWriteRunErrLog()
    glngNoteLength = GetNoteLength
    gblnHaveVersion = IsHaveVersion()
    gblnSameFTP = IsSampleFTP()
    Set grsErrRec = CopyNewRec(Nothing, True, , Array("类型", adInteger, 3, 0, "对象", adVarChar, 100, Empty, "信息", adVarChar, 1000, Empty))
    Call RecordErrMsg(MT_MsgHeader, "消息头", "操作:" & Decode(gotCurType, OT_OfficialUpgrade, "升级", OT_PreUpgrade, "预升", OT_Repair, "修复", OT_CheckFile, "收集") & " 开始:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
    gobjTrace.WriteInfo "GetBaseInfo", "工作站", gstrComputerName, "IP", gstrComputerIp, "System路径", gstrSystemPath, "临时目录", gstrTempPath, "记录运行日志", mblnWriteRunErrLog, "说明信息长度", glngNoteLength
    GetBaseInfo = True
    Exit Function
ErrH:
    strErrInfo = Err.Description
    gobjTrace.WriteInfo "GetBaseInfo", "基础信息获取发生严重错误", strErrInfo
    MsgBox "获取基础信息发生错误，请联系管理员！信息：" & vbNewLine & strErrInfo, vbInformation, App.Title
    Err.Clear
End Function

Public Sub RecordErrMsg(ByVal mtInput As MsgType, ByVal strErrObject As String, ByVal strErrInfo As String)
    Dim strSQL As String
    grsErrRec.AddNew Array("类型", "对象", "信息"), Array(mtInput, strErrObject, strErrInfo)
    If mtInput > MT_MsgHeader And mtInput < MT_MsgFoot Then
        On Error Resume Next
        '书写升级日志
        strSQL = "Zl_Zlclientupdatelog_Insert(" & SQLAdjust(strErrObject & ":" & strErrInfo) & "," & SQLAdjust(gstrTerminal) & ")"
        Call ExecuteProcedure(strSQL, "RecordErrMsg")
        If Err.Number <> 0 Then Err.Clear
        
        '书写运行日志
        If mblnWriteRunErrLog Then
            '类型=4 客户端升级错误
            '错误序号=0
            strSQL = "Zl_Zlerrorlog_Insert(" & SQLAdjust(gstrTerminal) & ",4,0," & SQLAdjust(strErrObject & ":" & strErrInfo) & "," & SQLAdjust(gstrAudsid) & " )"
            Call ExecuteProcedure(strSQL, "RecordErrMsg")
            If Err.Number <> 0 Then Err.Clear
        End If
    ElseIf mtInput = MT_MsgHeader Or mtInput = MT_MsgFoot Then
        On Error Resume Next
        
        '书写升级日志
        strSQL = "Zl_Zlclientupdatelog_Insert(" & SQLAdjust(strErrObject & ":" & strErrInfo) & "," & SQLAdjust(gstrTerminal) & ")"
        Call ExecuteProcedure(strSQL, "RecordErrMsg")
        If Err.Number <> 0 Then Err.Clear
    End If
End Sub

Public Function ReportInfo(ByVal strInfo As String)
    Dim strSQL      As String
    '书写升级日志
    On Error Resume Next
    strSQL = "Zl_Zlclientupdatelog_Insert(" & SQLAdjust(strInfo) & "," & SQLAdjust(gstrTerminal) & ",1)"
    Call ExecuteProcedure(strSQL, "ReportInfo")
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function SumErrMsg() As String
'功能：合并错误信息，产生信息汇总
    Dim strMsg As String, strPreObject As String
    
    grsErrRec.Filter = "类型=" & MT_MsgHeader
    If Not grsErrRec.EOF Then strMsg = grsErrRec!信息
    grsErrRec.Filter = "类型=" & MT_InitEnv
    Do While Not grsErrRec.EOF
        strMsg = strMsg & vbNewLine & grsErrRec!对象 & ":" & grsErrRec!信息
        grsErrRec.MoveNext
    Loop
    grsErrRec.Filter = "类型=" & MT_SvrConn
    Do While Not grsErrRec.EOF
        strMsg = strMsg & vbNewLine & grsErrRec!对象 & ":" & grsErrRec!信息
        grsErrRec.MoveNext
    Loop
    
    grsErrRec.Filter = "类型>" & MT_SvrConn & " And  类型<" & MT_ExeBat
    grsErrRec.Sort = "对象,类型"
    Do While Not grsErrRec.EOF
        If strPreObject <> grsErrRec!对象 Then
            strPreObject = grsErrRec!对象
            strMsg = strMsg & vbNewLine & grsErrRec!对象 & ":"
        End If
        strMsg = strMsg & vbNewLine & "  " & grsErrRec!信息
        grsErrRec.MoveNext
    Loop
    grsErrRec.Filter = "类型=" & MT_ExeBat
    Do While Not grsErrRec.EOF
        strMsg = strMsg & vbNewLine & grsErrRec!对象 & ":" & grsErrRec!信息
        grsErrRec.MoveNext
    Loop
    grsErrRec.Filter = "类型=" & MT_MsgFoot
    If Not grsErrRec.EOF Then strMsg = strMsg & vbNewLine & grsErrRec!信息
    SumErrMsg = strMsg
End Function

Private Function CheckFreeSpace() As Boolean
'功能：检查磁盘的剩余空间
    '检查磁盘空间，若少于1.5G,则提示不能预升级
    If gotCurType = OT_PreUpgrade Then
        If gobjFSO.Drives(Left(gstrSetupPath, 2)).FreeSpace / 1024 / 1024 < 500 Then
            gobjTrace.WriteInfo "磁盘空间检查", "信息", "空闲空间低于500MB,可能无法进行预升级"
            Call RecordErrMsg(MT_InitEnv, "磁盘空间检查", "空闲空间低于500MB,可能无法进行预升级")
        End If
    '正式升级或修复，至少要求200M空间
    Else
        If gobjFSO.Drives(Left(gstrSetupPath, 2)).FreeSpace / 1024 / 1024 < 200 Then
            gobjTrace.WriteInfo "磁盘空间检查", "信息", "空闲空间低于200MB,可能无法进行" & Decode(gotCurType, OT_OfficialUpgrade, "升级", OT_Repair, "修复", OT_CheckFile, "收集")
            Call RecordErrMsg(MT_InitEnv, "磁盘空间检查", "空闲空间低于200MB,可能无法进行" & Decode(gotCurType, OT_OfficialUpgrade, "升级", OT_Repair, "修复", OT_CheckFile, "收集"))
        End If
    End If
    CheckFreeSpace = True
End Function

Public Function GetActualPath(ByVal strSetupPath As String, ByVal ftFileType As FileType, ByVal strFile As String) As String
'功能：获取文件的实际路径
    Dim strKey As String, strPath As String
    
    If strSetupPath = "" Then
        Select Case ftFileType
            Case FT_Public
                strKey = "K_[PUBLIC]"
            Case FT_Apply
                strKey = "K_[APPSOFT]\APPLY"
            Case FT_Other, FT_AdditionFile
                strKey = "K_[APPSOFT]"
            Case FT_System
                strKey = "K_[SYSTEM]"
            Case FT_Help
                strKey = "K_[HELP]"
        End Select
    Else
        strKey = "K_" & strSetupPath
    End If
    strPath = gcllSetPath(strKey)
    GetActualPath = strPath & "\" & strFile
End Function

Public Function IsFileUpgade(ByVal strLoaclFile As String, ByVal strSvrVersion As String, ByVal strSvrModiTime As String, ByVal strSvrMD5 As String, Optional ByVal blnCheckReleated As Boolean)
'功能：是否更新下载
    Dim strlocVersion As String, strLocModiTime As String, strLocMd5 As String
    
    If Not gobjFSO.FileExists(strLoaclFile) Then
        '本地不存在，则判断服务器上是否存在，存在则升级，不存在则不升级
        IsFileUpgade = strSvrMD5 <> ""
        Exit Function
    End If
    '服务器文件不能存在，则不升级
    If strSvrMD5 = "" Then Exit Function
    '修改日期和版本不完整，则判断MD5
    If strSvrVersion = "" Or strSvrModiTime = "" Then
        strLocMd5 = FileMD5(strLoaclFile)
        IsFileUpgade = strLocMd5 <> strSvrMD5
    Else
        strlocVersion = GetCommpentVersion(strLoaclFile)
        If Len(strlocVersion) <> Len(strSvrVersion) And UCase(gobjFSO.GetFileName(strLoaclFile)) Like "ZL*" Then
            strLocMd5 = FileMD5(strLoaclFile)
            IsFileUpgade = strLocMd5 <> strSvrMD5
            Exit Function
        End If
        strLocModiTime = gobjFSO.GetFile(strLoaclFile).DateLastModified
        IsFileUpgade = strlocVersion <> strSvrVersion Or Format(strSvrModiTime, "YYYY-MM-DD hh:mm:ss") <> Format(strLocModiTime, "YYYY-MM-DD hh:mm:ss")
    End If
End Function

Public Function GetHisUpdateCommand(Optional ByVal blnOLD As Boolean) As String
'功能：获取自动升级的命令行
    Dim strCheck As String, strCommand As String
    Dim strUserName         As String, strPass      As String, strServer As String
    
    If blnOLD Then
        GetHisUpdateCommand = gstrConnectString & "||1||" & gstrAppEXE & "||" & gstrHisCommand & "||" & gstrHisInput
    ElseIf gstrCommand Like "ZLUPDATE:*" Then
        GetHisUpdateCommand = "ZLUPDATEEX:" & Cipher(gstrCommand)
    ElseIf gstrCommand Like "ZLUPDATEEX:*" Or gstrCommand Like "ZLUPDATENEW:*" Then
        GetHisUpdateCommand = gstrCommand
    Else
        GetHisUpdateCommand = "ZLUPDATENEW:" & Cipher(gstrCommand)
    End If
End Function

Public Sub ClearFolder(ByVal strFolder As String, Optional ByVal blnOk As Boolean)
'功能：清理执行文件夹
    Dim objFolder       As Folder, objFile          As File, objTmpFolder           As Folder
    Dim cllFolders      As New Collection, cllFiles       As New Collection
    Dim strTmpFile      As String, strTmpFloder As String
    Dim blnAdd          As Boolean
    Dim i               As Long
    On Error Resume Next
    If InStr(UCase(App.Path), "\ZLUPTMP") > 0 Or InStr(UCase(App.Path), "\APPLY") > 0 Then
        FileNormal gstrSetupPath & "\ZLHisCrust.EXE"
        Call gobjFSO.CopyFile(App.Path & "\ZLHisCrust.EXE", gstrSetupPath & "\ZLHisCrust.EXE", True)
        FileNormal App.Path & "\ZLHisCrust.EXE"
        Call gobjFSO.DeleteFile(App.Path & "\ZLHisCrust.EXE", True)
    End If
    If Err.Number <> 0 Then Err.Clear
    For Each objFolder In gobjFSO.GetFolder(strFolder).SubFolders
        '预升级不会删除预升级下载目录
        blnAdd = False
        If UCase(objFolder.Name) = "ZLPRETMP" Then
            If blnOk And (gotCurType = OT_OfficialUpgrade Or gotCurType = OT_Repair) Then
                blnAdd = True
            End If
        Else
            blnAdd = True
        End If
        If blnAdd Then
            cllFolders.Add objFolder.Path
            For Each objFile In objFolder.Files
                cllFiles.Add objFile.Path
            Next
        End If
    Next
    For i = 1 To cllFiles.Count
        Call gobjFSO.DeleteFile(cllFiles(i), True)
        If Err.Number <> 0 Then Err.Clear
    Next
    For i = 1 To cllFolders.Count
        Call gobjFSO.DeleteFolder(cllFolders(i), True)
        If Err.Number <> 0 Then Err.Clear
    Next
End Sub

Public Function FileNormal(ByVal strSource As String) As Boolean
'功能：将文件夹以及其子目录复制到另一个目录
    On Error Resume Next
    If gobjFSO.FileExists(strSource) Then
        If FileSystem.GetAttr(strSource) <> vbNormal Then
            FileSystem.SetAttr strSource, vbNormal
        End If
    End If
    
    FileNormal = Err.Number = 0
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function GetAdditionSetup(ByVal strFileName As String, ByVal strMD5 As String, ByVal strAdditionSetup As String) As String
'功能：获取附加安装路径，并对文件清理路径进行调整，清理文件路径不能包含附加安装路径中的路径
    Dim arrAdd      As Variant, i           As Integer, j       As Integer
    Dim arrTmp      As Variant, strLast     As String
    Dim arrAllPath  As Variant, strAllPath  As String, strTmp   As String
    Dim strAllFile  As String, strLocMd5    As String
    Dim strPath     As String
    
    If strAdditionSetup = "" Or strMD5 = "" Then Exit Function
    arrAdd = Split(UCase(strAdditionSetup), "|")
    For i = LBound(arrAdd) To UBound(arrAdd)
        arrTmp = Split(arrAdd(i), "\")
        strPath = ""
        If UBound(arrTmp) <> -1 Then
            If arrTmp(0) = "[APPSOFT]" Then
                strPath = gstrSetupPath
            ElseIf arrTmp(0) = "[PUBLIC]" Then
                If Not gobjFSO.FolderExists(gstrSetupPath & "\PUBLIC") Then
                    gobjFSO.CreateFolder (gstrSetupPath & "\PUBLIC")
                End If
                strPath = gstrSetupPath & "\PUBLIC"
            ElseIf arrTmp(0) = "[APPLY]" Then
                strPath = gstrSetupPath & "\APPLY"
            ElseIf arrTmp(0) = "[OS:]" Then '系统盘
                strPath = Left(gstrSystemPath, 2)
            ElseIf arrTmp(0) = "[X:]" Then '当前安装盘
                strPath = Left(gstrSetupPath, 2)
            End If
            If strPath <> "" Then
                strLast = Mid(arrAdd(i), Len(arrTmp(0) & "\") + 1)
                If strLast = "" Then
                    strTmp = strPath
                Else
                    strTmp = GetSubFloderByMach(strPath, strLast)
                End If
                If strTmp <> "" Then strAllPath = strAllPath & "|" & strTmp
            End If
        End If
    Next
    If strAllPath <> "" Then
        strAllPath = Mid(strAllPath, 2)
        arrAllPath = Split(strAllPath, "|")
        For i = LBound(arrAllPath) To UBound(arrAllPath)
            If gobjFSO.FileExists(arrAllPath(i) & "\" & strFileName) Then
                strLocMd5 = FileMD5(arrAllPath(i) & "\" & strFileName)
                If strMD5 <> strLocMd5 Then
                    strAllFile = strAllFile & "|" & arrAllPath(i) & "\" & strFileName
                    gobjTrace.WriteInfo "附加安装检测", "文件", arrAllPath(i) & "\" & strFileName, "信息", "该路径文件和服务器文件MD5不相同，需要附加安装"
                Else
                    gobjTrace.WriteInfo "附加安装检测", "文件", arrAllPath(i) & "\" & strFileName, "信息", "该路径文件和服务器文件MD5相同，不需要附加安装"
                End If
            Else
                strAllFile = strAllFile & "|" & arrAllPath(i) & "\" & strFileName
                gobjTrace.WriteInfo "附加安装检测", "文件", arrAllPath(i) & "\" & strFileName, "信息", "该路径存在但是文件不存在，因此需要下载并附加安装"
            End If
        Next
        If strAllFile <> "" Then strAllFile = Mid(strAllFile, 2)
    End If
    GetAdditionSetup = strAllFile
End Function

Private Function GetSubFloderByMach(ByVal strParentFloder As String, strMach As String) As String
'功能：获取匹配的自文件夹
'strParentFloder:父级文件夹
'strMach:匹配路径串
    Dim arrTmp      As Variant, strLast As String
    Dim objFolder   As Folder, blnLike  As Boolean, strLike As String
    Dim strTmp      As String, strReturn As String
    
    arrTmp = Split(strMach, "\")
    strLast = Mid(strMach, Len(arrTmp(0) & "\") + 1)
    If InStr(arrTmp(0), "[*]") > 0 Then
        strLike = Replace(arrTmp(0), "[*]", "*")
        For Each objFolder In gobjFSO.GetFolder(strParentFloder).SubFolders
            If UCase(objFolder.Name) Like strLike Then
                If strLast = "" Then
                    strTmp = strParentFloder & "\" & objFolder.Name
                Else
                    strTmp = GetSubFloderByMach(strParentFloder & "\" & objFolder.Name, strLast)
                End If
                If strTmp <> "" Then
                    strReturn = strReturn & "|" & strTmp
                End If
            End If
        Next
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        GetSubFloderByMach = strReturn
    Else
        If gobjFSO.FolderExists(strParentFloder & "\" & arrTmp(0)) Then
            If strLast = "" Then
                GetSubFloderByMach = strParentFloder & "\" & arrTmp(0)
            Else
                GetSubFloderByMach = GetSubFloderByMach(strParentFloder & "\" & arrTmp(0), strLast)
            End If
        End If
    End If
End Function

Public Function GetWrongFiles(ByVal strFileName As String, ByVal strSetupFile As String) As String
'功能：获取清理文件路径
    Dim varItem         As Variant, strFileTmp              As String
    Dim strWrongFile    As String
    
    For Each varItem In gcllSetPath
        strFileTmp = varItem & "\" & strFileName
        If UCase(strFileTmp) <> UCase(strSetupFile) Then
            If gobjFSO.FileExists(strFileTmp) Then
                If strWrongFile <> "" Then '处理[System]路径和[help]路径相同问题
                    If strWrongFile = "|" & strFileTmp Then
                    ElseIf InStr(strWrongFile, strFileTmp) = 0 Then
                        strWrongFile = strWrongFile & "|" & strFileTmp
                    End If
                Else
                    strWrongFile = strWrongFile & "|" & strFileTmp
                End If
            End If
        End If
    Next
    If strWrongFile <> "" Then strWrongFile = Mid(strWrongFile, 2)
    GetWrongFiles = strWrongFile
End Function

Private Function InitTerminal(ByRef strAudsid As String) As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrH
    strSQL = "Select Userenv('SessionID') Audsid ,Userenv('Terminal')  Terminal From dual"
    Set rsTmp = OpenSQLRecord(strSQL, "InitTerminal")
    
    If Not rsTmp.EOF Then
        strAudsid = rsTmp!Audsid
        InitTerminal = rsTmp!Terminal
    End If
    
    Exit Function
ErrH:
    MsgBox Err.Description
End Function

