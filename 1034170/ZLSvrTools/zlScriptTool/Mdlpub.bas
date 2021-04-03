Attribute VB_Name = "Mdlpub"
Option Explicit

Public Const HWND_TOPMOST = -1
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000

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
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字根类型...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

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

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_NT_WORKSTATION = 1
Private Const VER_NT_DOMAIN_CONTROLLER = 2
Private Const VER_NT_SERVER = 3
Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
    wServicePackMajor As Integer 'win2000 only
    wServicePackMinor As Integer 'win2000 only
    wSuiteMask As Integer 'win2000 only
    wProductType As Byte 'win2000 only
    wReserved As Byte
End Type
Private Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long

Private Const PRODUCT_UNLICENSED As Long = &HABCDABCD
Private Const PRODUCT_BUSINESS As Long = &H6
Private Const PRODUCT_BUSINESS_N As Long = &H10
Private Const PRODUCT_CLUSTER_SERVER As Long = &H12
Private Const PRODUCT_DATACENTER_SERVER As Long = &H8
Private Const PRODUCT_DATACENTER_SERVER_CORE As Long = &HC
Private Const PRODUCT_ENTERPRISE As Long = &H4
Private Const PRODUCT_ENTERPRISE_N As Long = &H1B
Private Const PRODUCT_ENTERPRISE_SERVER As Long = &HA
Private Const PRODUCT_ENTERPRISE_SERVER_CORE As Long = &HE
Private Const PRODUCT_ENTERPRISE_SERVER_IA64 As Long = &HF
Private Const PRODUCT_HOME_BASIC As Long = &H2
Private Const PRODUCT_HOME_BASIC_N As Long = &H5
Private Const PRODUCT_HOME_PREMIUM As Long = &H3
Private Const PRODUCT_HOME_PREMIUM_N As Long = &H1A
Private Const PRODUCT_HOME_SERVER As Long = &H13
Private Const PRODUCT_SERVER_FOR_SMALLBUSINESS As Long = &H18
Private Const PRODUCT_SMALLBUSINESS_SERVER As Long = &H9
Private Const PRODUCT_SMALLBUSINESS_SERVER_PREMIUM As Long = &H19
Private Const PRODUCT_STANDARD_SERVER As Long = &H7
Private Const PRODUCT_STANDARD_SERVER_CORE As Long = &HD
Private Const PRODUCT_STARTER As Long = &H8
Private Const PRODUCT_STORAGE_ENTERPRISE_SERVER As Long = &H17
Private Const PRODUCT_STORAGE_EXPRESS_SERVER As Long = &H14
Private Const PRODUCT_STORAGE_STANDARD_SERVER As Long = &H15
Private Const PRODUCT_STORAGE_WORKGROUP_SERVER As Long = &H16
Private Const PRODUCT_UNDEFINED As Long = &H0
Private Const PRODUCT_ULTIMATE As Long = &H1
Private Const PRODUCT_ULTIMATE_N As Long = &H1C
Private Const PRODUCT_WEB_SERVER As Long = &H11


Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function GetTempPath Lib "KERNEL32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Declare Function Htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long
Public Declare Function GetWindowsDirectory Lib "KERNEL32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

''''''''局域网登录
Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type
 
Public Const NO_ERROR = 0
Public Const CONNECT_UPDATE_PROFILE = &H1
' The following includes all the constants defined for NETRESOURCE,
' not just the ones used in this example.
Public Const RESOURCETYPE_DISK = &H1
Public Const RESOURCETYPE_PRINT = &H2
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCE_CONNECTED = &H1
Public Const RESOURCE_REMEMBERED = &H3
Public Const RESOURCE_GLOBALNET = &H2
Public Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEDISPLAYTYPE_SERVER = &H2
Public Const RESOURCEDISPLAYTYPE_SHARE = &H3
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const RESOURCEUSAGE_CONTAINER = &H2
' Error Constants:
Public Const ERROR_ACCESS_DENIED1 = 5&
Public Const ERROR_ALREADY_ASSIGNED = 85&
Public Const ERROR_BAD_DEV_TYPE = 66&
Public Const ERROR_BAD_DEVICE = 1200&
Public Const ERROR_BAD_NET_NAME = 67&
Public Const ERROR_BAD_PROFILE = 1206&
Public Const ERROR_BAD_PROVIDER = 1204&
Public Const ERROR_BUSY = 170&
Public Const ERROR_CANCELLED = 1223&
Public Const ERROR_CANNOT_OPEN_PROFILE = 1205&
Public Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&
Public Const ERROR_EXTENDED_ERROR = 1208&
Public Const ERROR_INVALID_PASSWORD = 86&
Public Const ERROR_NO_NET_OR_BAD_PATH = 1203&

Declare Function WNetAddConnection2 Lib "mpr.dll" Alias _
     "WNetAddConnection2A" (lpNetResource As NETRESOURCE, _
     ByVal lpPassword As String, ByVal lpUserName As String, _
     ByVal dwFlags As Long) As Long
      
Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias _
     "WNetCancelConnection2A" (ByVal lpName As String, _
     ByVal dwFlags As Long, ByVal fForce As Long) As Long
''''''''局域网登录

'---------------------------------------------------------------
Public gcnOracle As New ADODB.Connection    '数据库连接
Public Const gstrSysName = "中联软件"              '系统名称
Public gstrUserName As String               '用户名
Public gstr姓名 As String                   '用户姓名
Public gstr编号 As String                   '用户编号
Public gstrPassword As String               '用户口令
Public gstrServer As String                 '服务器名
Public gstrSql    As String                 '通用的SQL语句变量
Public rsTemp As New ADODB.Recordset        '公用临时recordset对象
Public gstrProductName As String

Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&

Public objFso As New FileSystemObject   '文件操作对象
Public gmobjCommon As Object ' zl9ComLib.clsComLib
Public blnExit As Boolean
Public gmdlOSType As Integer

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, ByRef objConn As ADODB.Connection) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strError As String
    
    On Error GoTo Hwanderr
    If objConn.State = adStateOpen Then objConn.Close
    
    With objConn
         .CursorLocation = adUseClient
         .ConnectionString = "Provider=MSDAORA.1;Password=" & strUserPwd & ";User ID=" & strUserName & ";Data Source=" & strServerName & ";Persist Security Info=True"
         .Open 'MSDAORA.1 要读取ORACLE blob字段 要使用 Provider= OraOLEDB.Oracle
    End With
    
    OraDataOpen = True
    Exit Function
Hwanderr:
   If err.Number <> 0 Then
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
        ElseIf InStr(strError, "ORA-ORA-03114") > 0 Then
            Exit Function
        Else
            MsgBox "由于用户、口令或服务器指定错误，无法注册。", vbInformation, gstrSysName
        End If
        OraDataOpen = False
        Exit Function
    End If
    OraDataOpen = False
End Function


Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    'strOld：原密码
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


Public Sub SelAll(objTxt As Control)
'功能：对文本框的的文本选中
    If TypeName(objTxt) = "TextBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Or InStr(strInput, """") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function


Public Function GetFileLineCount(ByVal txtStream As TextStream) As Long
    Do Until txtStream.AtEndOfStream
        txtStream.ReadLine
    Loop
    
    GetFileLineCount = txtStream.Line
End Function
Public Function LoadCustomPicture(strID As String) As StdPicture
'功能:将资源文件中的指定资源生成磁盘文件
'参数:ID=资源号,strExt=要生成文件的扩展名(如BMP)
'返回:生成文件名
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, "CUSTOM")
    intFile = FreeFile
    
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(Timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    Set LoadCustomPicture = VB.LoadPicture(strR)
    Kill strR
End Function

Public Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "", Optional gcnConnect As ADODB.Connection)
'功能：打开记录集
    If rsTemp.State = adStateOpen Then rsTemp.Close
'    rsTemp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle, adOpenStatic, adLockReadOnly
    If gcnConnect Is Nothing Then Set gcnConnect = gcnOracle
    rsTemp.CursorType = 1
    rsTemp.Open (IIf(strSQL = "", gstrSql, strSQL)), gcnConnect
    '  Call WriteLogFile(IIf(strSQL = "", gstrSql, strSQL))
    End Sub


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


Public Sub ApplyOEM_Picture(objPicture As Object, ByVal str属性 As String)
'针对各种图标应用OEM策略
    Dim strOEM As String
    Dim blnCorp As Boolean
    On Error Resume Next
    
    If gstrProductName <> "-" Then
        '处理状态栏图标的OEM策略
        If gstrProductName <> "中联" Then
            If Right(str属性, 1) = "B" Then
                '表示产品图片
                blnCorp = False
                str属性 = Mid(str属性, 1, Len(str属性) - 1)
            Else
                '表示公司徽标
                blnCorp = True
            End If
            
            strOEM = GetOEM(gstrProductName, blnCorp)
            If str属性 = "Picture" Then
                Set objPicture.Picture = LoadCustomPicture(strOEM)
            ElseIf str属性 = "Icon" Then
                Set objPicture.Icon = LoadCustomPicture(strOEM)
            End If
            
            If err <> 0 Then
                err.Clear
            End If
        End If
    End If
End Sub

Public Function GetOEM(ByVal strAsk As String, Optional ByVal blnCorp As Boolean = True) As String
    '-------------------------------------------------------------
    '功能：返回每个字线的ASCII码
    '参数：
    '返回：
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    'OEM图片有两种类型 ，一是指公司徽标，另一个是产品标识
    strCode = IIf(blnCorp = True, "OEM_", "PIC_")
    For intBit = 1 To Len(strAsk)
        '取每个字的ASCII码
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function

Public Sub ApplyOEM(objStatus As Object)
'针对状态栏应用OEM策略
    Dim strOEM As String
    On Error Resume Next
    
    If gstrProductName <> "-" Then
        objStatus.Panels(1).Text = gstrProductName & "软件"
        '处理状态栏图标的OEM策略
        If gstrProductName = "中联" Then
            Set objStatus.Panels(1).Picture = LoadCustomPicture("Logo")
        Else
            strOEM = GetOEM(gstrProductName)
            Set objStatus.Panels(1).Picture = LoadCustomPicture(strOEM)
            If err <> 0 Then
                err.Clear
                Set objStatus.Panels(1).Picture = LoadCustomPicture("Logo")
            End If
        End If
        objStatus.Panels(1).ToolTipText = ""
        objStatus.Height = 360
    End If
End Sub
Public Sub ExecuteProcedure(ByVal strCaption As String)
'功能：执行SQL语句
    gcnOracle.Execute gstrSql, , adCmdStoredProc
   ' Call WriteLogFile(gstrSql)

End Sub

Public Sub WriteLogFile(ByVal strLogContext As String)
'功能：将程序运行日志写入日志文件
'参数：strlogConText 日志内容
  Dim filename As String
  Dim logFile As TextStream
  
  filename = App.Path & "\" & Format(Now(), "yyyy-mm-dd") & "_" & "药品处理.log"
  '1、判断日志文件是否已经存在
  If objFso.FileExists(filename) Then
     Set logFile = objFso.OpenTextFile(filename, ForAppending)
  Else
     Set logFile = objFso.CreateTextFile(filename, True)
  End If
  
  '2、写入日志
  logFile.WriteLine (Format(Now(), "yyyy-mm-dd HH:MM:SS") & " ： ")
  logFile.WriteLine (strLogContext)
  logFile.Close
End Sub

Public Function ExcuteSql(ByVal Recordsettest As ADODB.Recordset, ByVal strSQL As String) As Boolean
'功能：16
'参数：16
    On Error GoTo ExcuteSqlError
    'Set Recordsettest = Nothing
    If Recordsettest.State = adStateOpen Then
       Recordsettest.Close
    End If
    'Recordset.CursorType = adOpenKeyset
    'Recordset.LockType = adLockOptimistic
    Call Recordsettest.Open(strSQL, gcnOracle, adOpenKeyset, adLockOptimistic, -1)
    ExcuteSql = True
    Exit Function
ExcuteSqlError:
    MsgBox "错误代码:" & err.Number & vbCrLf & _
            "错误描述:" & err.Description, vbCritical + vbOKOnly, "错误"
    ExcuteSql = False
End Function

Public Function ConnectUserPassword(ByVal RemoteName As String, ByVal LoginUser As String, ByVal LoginPassWord As String) As Boolean
     Dim NetR As NETRESOURCE
     Dim ErrInfo As Long
     Dim MyPass As String, MyUser As String
      
     NetR.dwScope = RESOURCE_GLOBALNET
     NetR.dwType = RESOURCETYPE_DISK
     NetR.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
     NetR.dwUsage = RESOURCEUSAGE_CONNECTABLE
     NetR.lpLocalName = "" '指定为本机映射盘符 如X:
     NetR.lpRemoteName = RemoteName '"\\192.168.0.3\常用软件"
     'NetR.lpComment = "Optional Comment"
     'NetR.lpProvider = ' Leave this undefined
     
     ErrInfo = WNetAddConnection2(NetR, LoginPassWord, LoginUser, CONNECT_UPDATE_PROFILE) '"snp1108$", "zlsoft\zq",
     If ErrInfo = NO_ERROR Then
        ConnectUserPassword = True
     Else
        ConnectUserPassword = False
     End If
End Function

Public Function GetWindowsVersion() As String
' 变量声明
Dim retOSVersionInf As OSVERSIONINFOEX
Dim retLng As Long

    '结构尺寸
    retOSVersionInf.dwOSVersionInfoSize = Len(retOSVersionInf)
    
    '获取 Windows 版本
    retLng = GetVersionEx(retOSVersionInf)
    
    If retLng = 0 Then
        GetWindowsVersion = "未知"
        Exit Function
    End If

    With retOSVersionInf
        If .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
            gmdlOSType = 0
        ElseIf .dwPlatformId = VER_PLATFORM_WIN32_NT And .dwMajorVersion = 5 Then
            gmdlOSType = 0
        ElseIf .dwPlatformId = VER_PLATFORM_WIN32_NT And .dwMajorVersion = 6 Then
            gmdlOSType = 1
        ElseIf .dwMajorVersion > 6 Then
            gmdlOSType = 1
        End If
        
        GetWindowsVersion = GetWindowsVersion & " [Version: " & .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber & "]"
    End With
End Function

Public Function IsPrivs(ByVal strPrivs As String, ByVal strPriv As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function


Public Function SearchPrintData(ByVal objVsf As Object, ByRef objPrintVsf As Object, Optional strNotPrintCol As String = "") As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strFormat As String
    Dim lngNotPrintCols As Long
    Dim lngPrintCol As Long
    
    If objPrintVsf.Cols = 0 Then Exit Function
    If strNotPrintCol <> "" Then
        lngNotPrintCols = UBound(Split(strNotPrintCol, ",")) + 1
        strNotPrintCol = "," & strNotPrintCol & ","
    End If
    
    objPrintVsf.Rows = objVsf.Rows
    objPrintVsf.FixedRows = objVsf.FixedRows
    
    objPrintVsf.Cols = 0
    lngPrintCol = -1
    For lngCol = 0 To objVsf.Cols - 1
        
        If objVsf.ColHidden(lngCol) = False And objVsf.TextMatrix(0, lngCol) <> "" Then
            
            If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                
                lngPrintCol = lngPrintCol + 1
                
                objPrintVsf.Cols = lngPrintCol + 1
                
                objPrintVsf.ColWidth(lngPrintCol) = objVsf.ColWidth(lngCol)
                objPrintVsf.ColAlignmentFixed(lngPrintCol) = objVsf.ColAlignment(lngCol)
                If objVsf.ColDataType(lngCol) = flexDTBoolean Then
                    objPrintVsf.ColAlignment(lngPrintCol) = 4
                Else
                    objPrintVsf.ColAlignment(lngPrintCol) = objVsf.ColAlignment(lngCol)
                End If
            End If
        End If
    Next
    
    If objPrintVsf.Cols = 0 Then Exit Function
    
    For lngRow = 0 To objVsf.Rows - 1

        objPrintVsf.RowHeight(lngRow) = IIf(objVsf.RowHeight(lngRow) < objVsf.RowHeightMin, objVsf.RowHeightMin, objVsf.RowHeight(lngRow))
        lngPrintCol = -1
        For lngCol = 0 To objVsf.Cols - 1
            
            If objVsf.ColHidden(lngCol) = False And objVsf.TextMatrix(0, lngCol) <> "" Then
                If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                
                    lngPrintCol = lngPrintCol + 1
                    
                    If objVsf.ColDataType(lngCol) = flexDTBoolean And lngRow >= objVsf.FixedRows Then
                        objPrintVsf.TextMatrix(lngRow, lngPrintCol) = IIf(Abs(Val(objVsf.TextMatrix(lngRow, lngCol))) = 1, "√", "")
                    Else
                        strFormat = objVsf.ColFormat(lngCol)
                        If strFormat = "" Then
                            objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Trim(objVsf.TextMatrix(lngRow, lngCol))
                        Else
                            objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Format(objVsf.TextMatrix(lngRow, lngCol), strFormat)
                        End If
                    End If
                End If
            End If
        Next
        Call SetMsfForeColor(objPrintVsf, lngRow, Val(objVsf.Cell(flexcpForeColor, lngRow, 1)))
    Next
    SearchPrintData = True
End Function

Public Sub SetMsfForeColor(ByRef msf As Object, ByVal lngRow As Long, ByVal lngColor As Long)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intCol As Integer

    With msf

        .Row = lngRow
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellForeColor = lngColor
        Next

    End With
End Sub
