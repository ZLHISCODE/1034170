Attribute VB_Name = "mdlMain"
Option Explicit

Public SplashObj As New frmSplash
Public gcnOracle As New ADODB.Connection    '公共数据库连接

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录

Public gstrUserFlag As String               '当前用户标志(两位表示)，第1位：是否DBA；第2位：系统所有者

Public gstrDbUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public gstrStation As String                '本工作站名称
Public gstrMenuSys As String                '系统菜单

Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstr单位名称 As String
Public glngSys As Long
Public gdtStart As Long
Public gblnEmerge As Boolean                '是否区分急诊 2008-12-24
Public gblnClearData As Boolean             '是否清空日志
Public gstr仪器设置 As String               '保存本机的仪器设置
Public Type T仪器设置
    ID      As Long
    类型    As Integer  '0-COM口方式 1-IP方式
    COM口   As Integer
    波特率  As Long
    数据位  As String
    校验位  As String
    停止位  As String
    握手    As String
    IP端口  As Long
    IP      As String
    主机     As Long
    字符模式 As String
    SaveAsID As Long
    编码名称 As String
    自动应答 As String  '自动应答间隔，单位秒，为<=0时不启用。
    可发已核标本 As Long '>0可以发 ,0<=不可以发
End Type
'-----------------------------------------
'发行码、注册码、发行码解析串、注册码解析串
Public gstrRegCode As String
Public gstrPublish As String
Public gstrParseRegCode As String
Public gstrParsePublish As String
'-----------------------------------------

Public gstrSystems As String

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'---------------------------------------------------------------
'-注册表 API 声明...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1

'---------------------------------------------------------------
'- 注册表 Api 常数...
'---------------------------------------------------------------
' Reg Data Types...
Const REG_SZ = 1                         ' Unicode空终结字符串
Const REG_EXPAND_SZ = 2                  ' Unicode空终结字符串
Const REG_DWORD = 4                      ' 32-bit 数字

' 注册表创建类型值...
Const REG_OPTION_NON_VOLATILE = 0       ' 当系统重新启动时，关键字被保留

' 注册表关键字安全选项...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字根类型...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004

' 返回值...
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- 注册表安全属性类型...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

'--- 公共部件
Public gobjComLib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gobjPrintMode As Object
Public g仪器() As T仪器设置


'---------------------------------------------------------------
'   授权、菜单、试用版本
'---------------------------------------------------------------
Public Sub Main()
    Dim lngReturn As Long
    Dim StrUnitName As String
    Dim BlnShowFlash As Boolean
    Dim strCode As String, IntCount As Integer, StrStyle As String
    Dim rsMenu As ADODB.Recordset, StrHaveSys As String
    
    
    
    '为实现XP风格，在显示窗体前必须执行该函数
    
    Call InitCommonControls
    '创建公共部件
    If gobjComLib Is Nothing Then Set gobjComLib = GetObject("", "zl9Comlib.clsComlib")
    If gobjCommFun Is Nothing Then Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    If gobjControl Is Nothing Then Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    If gobjDatabase Is Nothing Then Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    If gobjPrintMode Is Nothing Then Set gobjPrintMode = GetObject("", "zl9PrintMode.zlPrintMethod")
    
    BlnShowFlash = False
    Load SplashObj
    '由注册表中获取用户注册相关信息,如果用户单位名称不为空,则显示闪现窗体
    StrUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "提示", "")

    If StrUnitName <> "" Then
        With SplashObj
            '有两处需要处理
            Call gobjComLib.ApplyOEM_Picture(.ImgIndicate, "Picture")
            Call gobjComLib.ApplyOEM_Picture(.imgPic, "PictureB")
            .Show
            .lblGrant = StrUnitName
            StrUnitName = GetSetting("ZLSOFT", "注册信息", "开发商", "")
            If Trim(StrUnitName) = "" Then
                .Label3.Visible = False
                .lbl开发商.Visible = False
            Else
                .lbl开发商.Caption = ""
                For IntCount = 0 To UBound(Split(StrUnitName, ";"))
                    .lbl开发商.Caption = .lbl开发商.Caption & Split(StrUnitName, ";")(IntCount) & vbCrLf
                Next
            End If
            .LblProductName = GetSetting("ZLSOFT", "注册信息", "产品全称", "")
            .lbl技术支持商 = GetSetting("ZLSOFT", "注册信息", "技术支持商", "")
        End With
        
        BlnShowFlash = True
        DoEvents
    End If
    
    gstrStation = Space(200)
    lngReturn = GetComputerName(gstrStation, 200)
    gstrStation = Trim(gstrStation)
    If Len(gstrStation) > 1 Then
        gstrStation = Left(gstrStation, Len(gstrStation) - 1)
    Else
        gstrStation = "..."
    End If
    
    '用户注册
    frmUserLogin.Show 1
    If gcnOracle.State <> adStateOpen Then
        Unload frmUserLogin
        Unload SplashObj
        Exit Sub
    End If
    
    '初始化公共部件

    
    gobjComLib.InitCommon gcnOracle

    If gobjComLib.RegCheck = False Then
        Unload SplashObj
        Exit Sub
    End If
    
    '如果发行码无效（为空或为"-"），则退出
    gstrParsePublish = gobjComLib.zlRegInfo("产品简名")
    gstrParseRegCode = gobjComLib.zlRegInfo("单位名称", , -1)
    
    gstrSysName = gstrParsePublish & "软件"
    SaveSetting "ZLSOFT", "注册信息", "提示", gstrSysName
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrSysName"), gstrSysName
    gstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrVersion"), gstrVersion
    gstrAviPath = App.Path & "\附加文件"
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrAviPath"), gstrAviPath
    
    With SplashObj
        If BlnShowFlash = False Then
            .lblGrant = gstrParseRegCode
            .lbl技术支持商.Caption = gobjComLib.zlRegInfo("技术支持商", , -1)
            .LblProductName = gobjComLib.zlRegInfo("产品标题")
            
            strCode = gobjComLib.zlRegInfo("产品开发商", , -1)
            .lbl开发商.Caption = ""
            For IntCount = 0 To UBound(Split(strCode, ";"))
                .lbl开发商.Caption = .lbl开发商.Caption & Split(strCode, ";")(IntCount) & vbCrLf
            Next
            Call gobjComLib.ApplyOEM_Picture(.ImgIndicate, "Picture")
            .Show
            BlnShowFlash = True
        End If
        DoEvents
    End With
    
    '将用户注册相关信息写入注册表,供下次启动时显示
    SaveSetting "ZLSOFT", "注册信息", "单位名称", gstrParseRegCode
    SaveSetting "ZLSOFT", "注册信息", "产品全称", gobjComLib.zlRegInfo("产品标题")
    SaveSetting "ZLSOFT", "注册信息", "产品名称", gobjComLib.zlRegInfo("产品简名")
    SaveSetting "ZLSOFT", "注册信息", "技术支持商", gobjComLib.zlRegInfo("技术支持商", , -1)
    SaveSetting "ZLSOFT", "注册信息", "开发商", gobjComLib.zlRegInfo("产品开发商", , -1)
    SaveSetting "ZLSOFT", "注册信息", "WEB支持商简名", gobjComLib.zlRegInfo("支持商简名")
    SaveSetting "ZLSOFT", "注册信息", "WEB支持EMAIL", gobjComLib.zlRegInfo("支持商MAIL")
    SaveSetting "ZLSOFT", "注册信息", "WEB支持URL", gobjComLib.zlRegInfo("支持商URL")

    '多帐套、ZYB、2001-09-19修改
    '-------------------------------------------------------------
    '检查本机安装部件
    '-------------------------------------------------------------
    '-------------------------------------------------------------
    '调用帐套选择窗体
    '-------------------------------------------------------------
    gstrSystems = " (系统 =100 Or 系统 Is NULL)"
    
    '-------------------------------------------------------------
    '分析菜单及部件
    '-------------------------------------------------------------
'    Set rsMenu = MenuGranted
'    If rsMenu.EOF Then
'        MsgBox "您没有操作任何系统的权限,程序被迫退出！", vbInformation, gstrSysName
'        Unload SplashObj
'        Exit Sub
'    End If
    '-------------------------------------------------------------
    '创建同义词
    '-------------------------------------------------------------
    
    glngSys = 100
    Call CreateSynonyms(glngSys, 1208)
    
    gblnFromDB = IsFromDb
    
    If gblnFromDB Then
        gblnEmerge = gobjDatabase.GetPara("急诊标本", glngSys, 1208, 1)
    Else
        gblnEmerge = Val(GetSetting("ZLSOFT", "公共模块\zl9LisWork\frmLabMain", "急诊标本", 0))
    End If
    '-------------------------------------------------------------
    '选择调用不同风格导航台
    '-------------------------------------------------------------
    On Error Resume Next
    Err = 0
    
    Unload SplashObj
    
    CodeMan 1208
End Sub

Public Sub CodeMan(ByVal lngModul As Long)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '返回：
    '------------------------------------------------
    Dim clsPublic As New clsPublic
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstr单位名称 = gobjComLib.GetUnitName()
    Call GetUserInfo
    
    gstrPrivs = gobjComLib.GetPrivFunc(glngSys, lngModul)
    '-------------------------------------------------
    
    Select Case lngModul
        Case 1208
            clsPublic.InitClsPublic
    End Select
End Sub

Private Function CreateSynonyms(ByVal lngSys As Long, ByVal lngModul As Long)
    Dim strSQL As String
    '创建模块所需对象的同义词(如果已创建则不会再创建)
    On Error Resume Next
    strSQL = "Zl_Createsynonyms(" & lngSys & ")"
    gobjDatabase.ExecuteProcedure strSQL, "创建同义词"
End Function

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            Else
                MsgBox "由于用户、口令或服务器指定错误，无法注册。", vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    gstrDbUser = UCase(strUserName)
    gobjComLib.SetDbUser gstrDbUser
    OraDataOpen = True
    Exit Function
    
errHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function

Public Function OraDataClose() As Boolean
    '------------------------------------------------
    '功能： 关闭数据库
    '参数：
    '返回： 关闭数据库，返回True；失败，返回False
    '------------------------------------------------
    Err = 0
    On Error Resume Next
    gcnOracle.Close
    OraDataClose = True
    Err = 0

End Function

Public Function TranPasswd(strOld As String) As String
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

Public Function UpdatePassword(ByVal strUserName As String, ByVal strPasswd As String) As Boolean
    '-------------------------------------------------------------
    '功能：按人员ID，修改其密码
    '参数：CurrUser
    '      当前用户集
    '返回：如果成功则退回True，否则返回False
    '-------------------------------------------------------------
    Err = 0
    On Error GoTo ErrorHand
    
    DoEvents
    gcnOracle.Execute "alter user " & strUserName & " identified by " & strPasswd
    UpdatePassword = True
    Exit Function
    
ErrorHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    UpdatePassword = False

End Function

Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
'功能：写注册表
    Dim rc As Long                                      ' 返回代码
    Dim hKey As Long                                    ' 处理一个注册表关键字
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' 注册表安全类型
    
    lpAttr.nLength = 50                                 ' 设置安全属性为缺省值...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- 创建/打开注册表关键字...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' 创建/打开//KeyRoot//KeyName
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' 错误处理...
    
    '------------------------------------------------------------
    '- 创建/修改关键字值...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' 要让RegSetValueEx() 工作需要输入一个空格...
    
    ' 创建/修改关键字值
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))
                       
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' 错误处理
    '------------------------------------------------------------
    '- 关闭注册表关键字...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' 关闭关键字
    
    UpdateKey = True                                    ' 返回成功
    Exit Function                                       ' 退出
CreateKeyError:
    UpdateKey = False                                   ' 设置错误返回代码
    rc = RegCloseKey(hKey)                              ' 试图关闭关键字
End Function

'-------------------------------------------------------------------------------------------------
'sample usage - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'-------------------------------------------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'功能：读注册表
    Dim i As Long                                           ' 循环计数器
    Dim rc As Long                                          ' 返回代码
    Dim hKey As Long                                        ' 处理打开的注册表关键字
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' 注册表关键字数据类型
    Dim tmpVal As String                                    ' 注册表关键字的临时存储器
    Dim KeyValSize As Long                                  ' 注册表关键字变量尺寸
    
    ' 在 KeyRoot {HKEY_LOCAL_MACHINE...} 下打开注册表关键字
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 打开注册表关键字
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误...
    
    tmpVal = String$(1024, 0)                             ' 分配变量空间
    KeyValSize = 1024                                       ' 标记变量尺寸
    
    '------------------------------------------------------------
    ' 检索注册表关键字的值...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' 获得/创建关键字的值
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 错误处理
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' 决定关键字值的转换类型...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' 搜索数据类型...
    Case REG_SZ, REG_EXPAND_SZ                              ' 字符串注册表关键字数据类型
        sKeyVal = tmpVal                                     ' 复制字符串的值
    Case REG_DWORD                                          ' 四字节注册表关键字数据类型
        For i = Len(tmpVal) To 1 Step -1                    ' 转换每一位
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 一个字符一个字符地生成值。
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' 转换四字节为字符串
    End Select
    
    GetKeyValue = sKeyVal                                   ' 返回值
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
    Exit Function                                           ' 退出
    
GetKeyError:    ' 错误发生过后进行清除...
    GetKeyValue = vbNullString                              ' 设置返回值为空字符串
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
End Function

Public Sub CheckDBConnect()
    On Error GoTo ConnErr
    If gcnOracle.State <> 1 Then gcnOracle.Open
    gcnOracle.Execute "select '测试'  from dual"
    Exit Sub
ConnErr:
    On Error Resume Next
    If gcnOracle.State = 1 Then
        gcnOracle.Close
    End If
End Sub
Public Sub GetUserInfo()
'功能:得到用户的信息

    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String
    
    rsTemp.CursorLocation = adUseClient
    On Error GoTo errHand
    
    With rsTemp
        strSQL = "select P.*,D.编码 as 部门编码,D.名称 as 部门名称,M.部门ID" & _
                " from 上机人员表 U,人员表 P,部门表 D,部门人员 M " & _
                " Where U.人员id = P.id And P.ID=M.人员ID and  M.缺省=1 and M.部门id = D.id and U.用户名=user"
        .Open strSQL, gcnOracle, adOpenKeyset
                
        If .RecordCount <> 0 Then
            glngUserId = .Fields("ID").Value                '当前用户id
            gstrUserCode = .Fields("编号").Value            '当前用户编码
            gstrUserName = .Fields("姓名").Value            '当前用户姓名
            gstrUserAbbr = IIf(IsNull(.Fields("简码").Value), "", .Fields("简码").Value)          '当前用户简码
            glngDeptId = .Fields("部门id").Value            '当前用户部门id
            gstrDeptCode = .Fields("部门编码").Value        '当前用户
            gstrDeptName = .Fields("部门名称").Value        '当前用户
        Else
            glngUserId = 0
            gstrUserCode = ""
            gstrUserName = ""
            gstrUserAbbr = ""
            glngDeptId = 0
            gstrDeptCode = ""
            gstrDeptName = ""
        End If
        .Close
    End With
    Exit Sub
errHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
    Err = 0
End Sub


Private Function IsFromDb() As Boolean
    '是否从数据库读取参数
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strSet As String
    
    Dim aPorts As Variant, i As Integer, lngID As Long
    On Error GoTo errH
    
    strSQL = "Select 编号 as 系统 From zlSystems Where Trunc(编号/100)=1 And 版本号 Like '10.24.%'"
    Set rsTmp = gcnOracle.Execute(strSQL)
    Do Until rsTmp.EOF
        IsFromDb = True
        rsTmp.MoveNext
    Loop
    
    If IsFromDb Then
        '检查系统是否有参数，如果没有，从本机的注册表中读。
        
        strSet = Trim(gobjDatabase.GetPara("本机连接仪器", glngSys, 1208, ""))
        If strSet = "" Then
            Err = 0: On Error Resume Next
            aPorts = GetAllSettings("ZLSOFT", "公共模块\ZlLISSrv")
            On Error GoTo errH
            
            If Not IsEmpty(aPorts) Then
                ReDim g仪器(UBound(aPorts))
                
                For i = LBound(aPorts) To UBound(aPorts)
                    lngID = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Device", 0))
                    If lngID > 0 Then
                        If aPorts(i, 0) Like "COM*" Then
                            g仪器(i).类型 = 0
                            g仪器(i).COM口 = Val(Replace(aPorts(i, 0), "COM", ""))

                            g仪器(i).字符模式 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "InputMode", "0"))
                        Else
                            g仪器(i).类型 = 1
                            g仪器(i).COM口 = 0
                            g仪器(i).字符模式 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "InMode", "0"))
                        End If
                        
                        With g仪器(i)
                            .ID = lngID
                            g仪器(i).波特率 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Speed", "9600"))
                            g仪器(i).数据位 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "DataBit", "8"))
                            g仪器(i).校验位 = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Parity", "N")
                            g仪器(i).停止位 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "StopBit", "1"))
                            g仪器(i).握手 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "HandShaking", "0"))
                            .IP端口 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Port", "6666"))
                            .IP = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "IP", "127.0.0.1")
                            .SaveAsID = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "SaveAs", "0"))
                            .主机 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Host", "0"))
                            
                            .自动应答 = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Auto", "0")
                            .可发已核标本 = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "blnSend", "1")
                        End With
                    End If
                Next
                
                gblnFromDB = True
                Call SavePortsSetting
            End If
        End If
    End If
    
    Exit Function
errH:
End Function
