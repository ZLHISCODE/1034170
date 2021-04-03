Attribute VB_Name = "mdlPublic"
Option Explicit

'调用参数： {+}1405{+}ZLHIS[+]ZLHIS[+]HIS[+]0{+}false{+}false{+}0{+}0{+}false

Public gstrLogPath As String        '日志文件
Public gstrImages As String         '消息参数 strImages
Public glngOrderID As Long          '消息参数 lngOrderID
Public gstrDBConnection As String   '消息参数 strDBConnection
Public gblnMoved As Boolean         '消息参数 blnMoved
Public gbAdd As Boolean             '消息参数 bAdd
Public gintImageInterval As Integer '消息参数 intImageInterval
Public glngSys As Long              '消息参数 lngSys
Public gblnReconnectDB As Boolean   '消息参数 blnReconnectDB
Public gstrZLHIS主机字符串 As String '消息参数 strDBServer
Public gstr用户名 As String          '消息参数 strDBUser
Public gstr密码 As String            '消息参数 strDBPassword
Public gbln是否转换密码 As Boolean '消息参数 blnTransPassword
Public gfrmViewImage As frmViewImage    '消息循环的主窗体
Public gobjPacsCore As Object       '观片对象
Public glngPreWndProc As Long       '原来的消息处理程序
Public glngLog As Long              '是否记录日志；0---参数未赋值；1---记录日志；2---不记录日志

Public Const HIS_CAPTION = "中联影像观片窗口"
Public Const MSG_SPLIT = "{+}"

Private mobjRegister As Object                  '10.35.10之后的注册对象
Public glngModule As Long                       '模块号
Public gblnBefore3510 As Boolean                '区分10.35.10前后版本。True=10.35.10之前版本,不使用zlRegister，初始化comlib时需要SetDbUser和RegCheck
Public gzlComLib As Object                      '公共数据库处理模块zlComLib
Public gcnOracle As ADODB.Connection            '公共数据库连接

Public Const gstrSysName As String = "影像观片"

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Enum LogType
    ltError = 0
    ltDebug = 1
End Enum

Public Function errHandle(errSubName As String, errTitle As String, Optional errDesc As String = "") As Long
'------------------------------------------------
'功能：错误处理
'参数： logSubName  --  产生错误的函数名
'       logTitle   -- 错误名称
'       logDesc   --  错误描述
'返回：1-程序继续Resume；0-程序退出
'------------------------------------------------
    
    errHandle = 0
    
    '记录错误日志
    Call WriteCommLog("zlSoftViewImage,错误--" & errSubName, errTitle & "，错误代码= " & err.Number, errDesc & "，错误描述=" & err.Description, ltError)
    
    '提示错误
    MsgBox errTitle & errDesc, vbOKOnly, "观片接口zlSoftViewImage出现错误"
    
    '清除错误
    err.Clear
    
End Function

Public Sub WriteCommLog(logSubName As String, logTitle As String, logDesc As String, ByVal ltLogType As LogType)
'------------------------------------------------
'功能：记录通讯日志
'参数： logSubName  --  产生日志的函数名
'       logTitle   -- 日志名称
'       logDesc   --  日志内容
'       ltLogType --  日志类型
'返回：无
'------------------------------------------------
    Dim strLog As String
    Dim strFileName As String
    
    On Error GoTo err
    
    If glngLog = 0 Then
        glngLog = Val(GetSetting("ZLSOFT", "公共模块\zl9PacsCore\zlSoftViewImage\", "Log", 2))
    End If
    
    'Log=1，才记录日志
    If glngLog <> 1 And ltLogType <> ltError Then Exit Sub
    
    strFileName = gstrLogPath & "\Interface" & Format(Date, "YYYY-MM-DD") & ".log"
    
    strLog = Now() & " 标题： " & logTitle & vbCrLf & "   函数： " & logSubName & vbCrLf & "   日志内容：" & logDesc & vbCrLf
    
    '错误日志增加标记，方便查看分析
    If ltLogType = ltError Then
        strLog = "▲▲错误▲▲：" & strLog
    End If
    
    Open strFileName For Append As #1
    Print #1, strLog
    Close #1
    
    Exit Sub
err:
    Close #1
End Sub

Public Function GetLogDir() As String
'------------------------------------------------
'功能：获取日志目录，如果目录不存在，则创建目录
'参数：无
'返回：日志所在目录
'------------------------------------------------
    Dim strLogPath As String
    Dim strBackupPath As String
    
    On Error GoTo err
    
    strLogPath = Mid(App.Path, 1, InStr(5, App.Path, "\"))
    strLogPath = strLogPath & "Log\日志跟踪\100_PACS观片日志"
    
    
    Call MkLocalDir(strLogPath + "\")
    
    GetLogDir = strLogPath
   
    Exit Function
err:
    GetLogDir = App.Path & "\100_PACS观片日志"
    Call MkLocalDir(GetLogDir + "\")
End Function

Public Function ProcessMessage(strMsg As String) As Long
'------------------------------------------------
'功能：处理接收到的消息
'参数：strMsg -- 调用exe时传入的参数串
'返回：无
'------------------------------------------------
    
    Dim lngPartType As Long
    Dim strDBUser As String
    Dim lngPatientID As Long
    Dim lngClinicID As Long
    Dim lngDeptID As Long
    Dim lngOrderID As Long
    
    On Error GoTo err
    ProcessMessage = 1
    
    '传入的参数定义，参数的连接符是三个字符“{+}”
    '参数格式：strImages{+}lngOrderID{+}strDBConnection{+}blnMoved{+}bAdd{+}intImageInterval{+}lngSys{+}blnReconnectDB
    '参数解释： strImages --- 图象号,规则是“序列UID1|1-3;5-27;33-100+序列UID2|全部”,全部表示打开全部图象
    '           lngOrderID --- 医嘱ID
    '           strDBConnection --- 数据库连接串，包含“服务名[+]用户名[+]密码[+]密码是否转换”，连接符是三个字符“[+]”
    '                          当“密码”是用户登录密码时，“密码是否转换”=1；当“密码”是数据库登录密码时，“密码是否转换”=0
    '           blnMoved --- 数据是否被转储
    '           bAdd --- 可选参数，默认值False，新图像是增加进观片站，还是替换原观片站所有图像，True为增加，Fasle为替换
    '           intImageInterval --- 可选参数，默认值0，打开图像的间隔，只对打开全部序列,且序列中图像数量>100时有效
    '           lngSys --- 可选参数，默认,100，系统序号
    '           blnReconnectDB --- 可选参数，默认值False，是否重新连接数据库。第一次打开观片时自动连接数据库，之后再打开观片，
    '                           由blnReconnectDB参数决定是否重新连接数据库。
    '                           =True，使用strDBConnection参数重新连接数据库；=False，不再重新连接数据库，使用观片部件现在的数据库连接
    '
    
    '先处理固定参数
    If UBound(Split(strMsg, MSG_SPLIT)) >= 3 Then
        gstrImages = Split(strMsg, MSG_SPLIT)(0)
        glngOrderID = Val(Split(strMsg, MSG_SPLIT)(1))
        gstrDBConnection = Split(strMsg, MSG_SPLIT)(2)
        gblnMoved = (UCase(Split(strMsg, MSG_SPLIT)(3)) = "TRUE")
    Else
        Call WriteCommLog("错误--zlSoftShowHisForms.ProcessMessage", "解析参数", "解析参数出错，参数数量不够4个，参数为：" & strMsg, ltError)
        Exit Function
    End If
    
    '再处理可选参数
    If UBound(Split(strMsg, MSG_SPLIT)) >= 4 Then
        gbAdd = (UCase(Split(strMsg, MSG_SPLIT)(4)) = "TRUE")
    Else
        gbAdd = False
    End If
    
    If UBound(Split(strMsg, MSG_SPLIT)) >= 5 Then
        gintImageInterval = Val(Split(strMsg, MSG_SPLIT)(5))
    Else
        gintImageInterval = 0
    End If
    
    If UBound(Split(strMsg, MSG_SPLIT)) >= 6 Then
        glngSys = Val(Split(strMsg, MSG_SPLIT)(6))
    Else
        glngSys = 100
    End If
    
    If UBound(Split(strMsg, MSG_SPLIT)) = 7 Then
        gblnReconnectDB = (UCase(Split(strMsg, MSG_SPLIT)(7)) = "TRUE")
    Else
        gblnReconnectDB = False
    End If
    
    If CreatePacsCore = False Then
        Exit Function
    End If
    
    Call WriteCommLog("zlSoftShowHisForms.ProcessMessage", "调用观片", "观片的参数是：gstrImages=" & gstrImages & ",glngOrderID=" & glngOrderID _
        & ",gstrDBConnection=" & gstrDBConnection & ",gblnMoved=" & gblnMoved & ",gbAdd=" & gbAdd & ",gintImageInterval=" & gintImageInterval _
        & ",glngSys=" & glngSys & ",gblnReconnectDB=" & gblnReconnectDB, ltDebug)
    
    Call gobjPacsCore.CallOpenViewer(gstrImages, glngOrderID, Nothing, gcnOracle, gblnMoved, gbAdd, gintImageInterval, glngSys)
    
    ProcessMessage = 0
    Exit Function
    
err:
    Call WriteCommLog("错误--zlSoftShowHisForms.ProcessMessage", "处理接收到的消息，出现错误，收到的消息是：" & strMsg & "，错误代码= " & err.Number, "，错误描述=" & err.Description, ltError)
End Function

'******************************************************************************************************************
'功能：创建PACS观片对象
'参数：无
'返回：创建成功,返回true,否则返回False
'说明：
'******************************************************************************************************************
Private Function CreatePacsCore() As Boolean

    err = 0: On Error Resume Next
    If Not gobjPacsCore Is Nothing Then CreatePacsCore = True: Exit Function
    
    Set gobjPacsCore = CreateObject("zl9PacsCore.clsViewer")
    
    If err <> 0 Then
        MsgBox "未找到 zl9PacsCore 部件，可能是程序版本不支持，请检查该站点是否部署了此部件!", vbInformation + vbOKOnly, "提示信息"
        Exit Function
    End If
    
    CreatePacsCore = True
    
End Function

Public Function CloseAllForms() As Boolean

    On Error GoTo err
    
    '关闭消息循环主窗口
    If Not gfrmViewImage Is Nothing Then
        Unload gfrmViewImage
        Set gfrmViewImage = Nothing
    End If
    
    CloseAllForms = True
    
    Exit Function
err:
    Call WriteCommLog("错误--zlSoftViewImage.CloseAllForms", "退出程序，关闭所有窗口，出现错误，错误代码= " & err.Number, "，错误描述=" & err.Description, ltError)
    Resume Next
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'功能：创建本地目录
'参数： strDir－－本地目录
'返回：无
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function InitInterface(ByVal strDBUser As String) As Boolean
'------------------------------------------------
'功能：初始化接口，创建ComLib，连接数据库
'参数：无
'返回：True-成功；False-失败
'------------------------------------------------
    
    On Error GoTo err
    InitInterface = False
    
    '初始化系统号为100，模块号为1289
    glngSys = 100
    glngModule = 1289
        
On Error Resume Next
    If mobjRegister Is Nothing Then
        Set mobjRegister = GetObject("", "zlRegister.clsRegister")
        If mobjRegister Is Nothing Then gblnBefore3510 = True '35.10之前的版本
    End If
    
    err.Clear
On Error GoTo err
    If gzlComLib Is Nothing Then
        If gblnBefore3510 Then
            '10.35.10之前的版本
            Set gzlComLib = CreateObject("zl9ComLib.clsComLib")
        Else
            '10.35.10之后的版本
            Set gzlComLib = GetObject("", "zl9ComLib.clsComLib")
        End If
    End If
    
    '如果是从RIS启动的DLL，数据库连接gzlComLib.CurrentConn是空的，需要从注册表读取用户名密码，并且连接数据库
    If gzlComLib.CurrentConn Is Nothing Then
        '从注册表读取用户名密码，连接数据库
        
        '如果gcnOracle不存在，要新建一个
        If gcnOracle Is Nothing Or gblnReconnectDB = True Then
            Set gcnOracle = New ADODB.Connection
            Call ConnectDB(strDBUser)
        End If

        '初始化公共部件
        gzlComLib.InitCommon gcnOracle
        
        If gblnBefore3510 = True Then
            '10.35.10之前的版本
            If gzlComLib.RegCheck = False Then
                
                Exit Function
            End If
        End If
    Else
        '如果是从HIS导航台启动的DLL，则创建zl9ComLib之后，会自动包含有gzlComLib.CurrentConn
        '现在暂时没有从 CodeMan中取得 gcnOracle，所以需要从zl9ComLib取得gcnOracle对象
        
        If gcnOracle Is Nothing Then Set gcnOracle = gzlComLib.CurrentConn
    End If
    
    InitInterface = True
    
  
    Exit Function
err:
    If errHandle("zlSoftShowHisForms.InitInterface", "初始化接口出错", err.Description) = 1 Then Resume
End Function

Public Function ConnectDB(ByVal strDBUser As String) As Boolean
'------------------------------------------------
'功能：连接数据库，从注册表中读取加密后的数据库连接信息：用户名，密码，服务名
'参数：
'返回：True-成功；False-失败
'------------------------------------------------
    Dim strDBPassword As String
    Dim strDBServer As String
    Dim blnTransPassword As Boolean
    
    ConnectDB = False
    
    On Error GoTo err
    
    If gcnOracle.State <> adStateOpen Then
        strDBServer = gstrZLHIS主机字符串
        strDBUser = gstr用户名
        strDBPassword = gstr密码
        blnTransPassword = gbln是否转换密码
                
        '连接数据库
        If OraDataOpen(strDBServer, strDBUser, strDBPassword, blnTransPassword) = False Then
           
            Exit Function
        End If
    End If
    
    ConnectDB = True
    Exit Function
err:
    If errHandle("zlSoftViewImage.ConnectDB", "连接数据库函数出现错误", err.Description) = 1 Then Resume
End Function

Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, ByVal blnTransPassword As Boolean) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '   blnTransPassword ： 是否需要转换密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strError As String
    
    On Error GoTo ErrHand
    

    If gblnBefore3510 = True Then
        '如果是10.35.10之前的版本，直接用用户名和密码登录数据库
        OraDataOpen = OpenOracle(gcnOracle, strServerName, strUserName, IIf(UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM", strUserPwd, IIf(blnTransPassword = True, TranPasswd(strUserPwd), strUserPwd)))
    Else
        '如果是10.35.10之后的版本，使用zlRegister获取数据库连接
        Set gcnOracle = mobjRegister.GetConnection(strServerName, strUserName, strUserPwd, blnTransPassword, , strError, True)
        If gcnOracle.State = adStateOpen Then
            OraDataOpen = True
        Else
            OraDataOpen = False
        End If
    End If
    
    If OraDataOpen = True Then
        strUserName = UCase(strUserName) '这里为什么要强制大写？是不是comlib的要求？
        If gblnBefore3510 = True Then
            '10.35.10之前的版本
            gzlComLib.SetDbUser strUserName
        End If
    End If
    
    Exit Function
    
ErrHand:
    
    If errHandle("zlSoftViewImage.OraDataOpen", "连接数据库出错", err.Description) = 1 Then Resume
    OraDataOpen = False
End Function

Private Function OpenOracle(ByRef cnOrcle As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的Oracle数据库
    '参数：
    '   cnOrcle ：数据库连接
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strError As String
    
    On Error Resume Next
    err = 0
    DoEvents
    With cnOrcle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If err <> 0 Then
            '保存错误信息
            strError = err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            OpenOracle = False
            Exit Function
        End If
    End With
    
    OpenOracle = True
    err = 0
    
    Exit Function
    
End Function

Private Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    
    strNew = ""
    
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    
    TranPasswd = strNew
End Function

