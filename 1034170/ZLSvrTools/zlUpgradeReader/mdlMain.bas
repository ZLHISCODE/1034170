Attribute VB_Name = "mdlMain"
Option Explicit
Public gstrDBUser As String
Public gcnOracle As ADODB.Connection
Public gblnOwner As Boolean
Public gstrSysname As String '程序名称

Public gstrSystems As String '系统名称
Public gstr用户单位名称 As String '已登录时不为空

Public mclsAppTool As New zl9AppTool.clsAppTool

Public rsMenu As ADODB.Recordset
Public rsMenuPEIS As ADODB.Recordset

'-------------------------------------------------------------
Public Const GWL_EXSTYLE = (-20)
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'-------------------------------------------------------------
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'-------------------------------------------------------------
' Reg Data Types...
Const REG_SZ = 1                         ' Unicode空终结字符串
Const REG_EXPAND_SZ = 2                  ' Unicode空终结字符串
Const REG_DWORD = 4                      ' 32-bit 数字

Public Const WinStyle = &H40000

' 注册表关键字安全选项...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字根类型...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_USERS = &H80000003

' 返回值...
Public Const ERROR_SUCCESS = 0
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

'---读写INI文件的API声明
#If Win32 Then
   Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal Filename As String) As Integer
#Else
   Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal Filename As String) As Integer
#End If
'----------------------

Public Enum 工具清单
    导航功能清单 = 10
    字典管理工具 = 11
    消息收发工具 = 12
    系统选项设置 = 13
    EXCEL报表工具 = 14
    本地参数管理 = 15
End Enum

Public Sub Main()
    
    Call InitCommonControls
    
    gblnOwner = False
    gstrDBUser = ""
    gstrSysname = "升级说明阅读器"
    gstr用户单位名称 = ""
    
    '用户注册
    frmUserLogin.Show 1
    If gcnOracle Is Nothing Then
        Set gcnOracle = New ADODB.Connection
    End If
    
    If gcnOracle.State = adStateOpen Then
        '初始化公共部件
        InitCommon gcnOracle
        
        If RegCheck = False Then
            Exit Sub
        End If
        
        '-------------------------------------------------------------
        '版本检查
        '-------------------------------------------------------------
        Select Case zlRegInfo("授权性质")
            Case "1"
                '正式
                SaveSetting "ZLSOFT", "注册信息", "Kind", ""
            Case "2"
                '试用
                SaveSetting "ZLSOFT", "注册信息", "Kind", "试用"
            Case "3"
                '测试
                SaveSetting "ZLSOFT", "注册信息", "Kind", "测试"
            Case Else
                '不对
                MsgBox "授权性质不正确，程序被迫退出！", vbInformation, gstrSysname
                Exit Sub
        End Select
    
        '多帐套、ZYB、2001-09-19修改
        '-------------------------------------------------------------
        '检查本机安装部件
        '-------------------------------------------------------------
        If TestComponent = False Then
            MsgBox "您没有操作任何系统的权限，程序被迫退出！", vbInformation, gstrSysname
            Exit Sub
        End If
        
        '-------------------------------------------------------------
        '调用帐套选择窗体
        '-------------------------------------------------------------
        With FrmAccoutChoose
            gstrSystems = .Show_me
            If .BlnSelect = False Then
                Exit Sub
            End If

            If gstrSystems = "" Then
                MsgBox "您没有操作任何系统的权限，程序被迫退出！", vbInformation, gstrSysname
                Exit Sub
            End If
            
            If gstrSystems <> "REPORT" Then
                gstrSystems = " 系统 in (" & gstrSystems & ")"
            End If
        End With
        
        '-------------------------------------------------------------
        '分析菜单及部件
        '-------------------------------------------------------------
        
        Set rsMenu = MenuGranted("")
        Set rsMenuPEIS = MenuGranted("PEIS")
        
        If rsMenu.EOF Then
            MsgBox "您没有操作任何系统的权限,程序被迫退出！", vbInformation, gstrSysname
            Exit Sub
        End If
        
        gstr用户单位名称 = zlRegInfo("单位名称", , -1)
        
        Call frmMain.Show_me(1) '0- 未登录方式 1－已登录方式
    Else
        Call frmMain.Show_me(0) '0- 未登录方式 1－已登录方式
    End If
    
    
End Sub

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Integer
    
    Err.Clear: On Error Resume Next
    DoEvents
    
    If gcnOracle Is Nothing Then
        Set gcnOracle = New ADODB.Connection
    End If
    With gcnOracle
        If .State = 1 Then .Close
        
        '.Provider = "MSDataShape"
        '.Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        
        .CursorLocation = adUseClient
        .Provider = "OraOLEDB.Oracle"
        .Open strServerName, strUserName, strUserPwd
        
        If Err <> 0 Then
            MsgBox "连接失败！（请确保用户名与密码输入无误）", vbInformation, App.Title
            Err.Clear: Exit Function
        End If
    End With
        
    '是否所有者用户
    If UCase(strUserName) <> "SYS" And UCase(strUserName) <> "SYSTEM" Then
        strSql = "Select 1 From zlSystems Where 所有者=USER"
        Set rsTmp = New ADODB.Recordset
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, gcnOracle, adOpenKeyset, adLockReadOnly
        gblnOwner = Not rsTmp.EOF
    End If
    
    gstrDBUser = strUserName
    
    OraDataOpen = True
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


Private Function TestComponent() As Boolean
    '如果没有任何部件可使用，则返回假
    TestComponent = False
    
    Dim strObjs As String, strCodes As String, strSql As String
    Dim objComponent As Object
    Dim resComponent As New ADODB.Recordset
    
    '--由注册表获取授权部件--
    strObjs = GetSetting("ZLSOFT", "注册信息", "本机部件", "")
    If strObjs <> "" Then TestComponent = True: Exit Function
    '--分析授权安装部件--
    With resComponent
        strSql = "Select Distinct Upper(g.部件) As 部件" & vbCrLf & _
                " From zlPrograms g, zlRegFunc r" & vbCrLf & _
                " Where g.序号 = r.序号 And Trunc(g.系统 / 100) = r.系统 And Upper(g.部件) <> 'ZL9REPORT'"
        
        If .State = adStateOpen Then .Close
        .Open strSql, gcnOracle, adOpenKeyset
        Err = 0: On Error Resume Next
        Do While Not .EOF
            Err = 0
            Set objComponent = CreateObject(!部件 & ".Cls" & Mid(!部件, 4))
            If Err = 0 Then strObjs = strObjs & IIf(strObjs = "", "", ",") & "'" & !部件 & "'"
            .MoveNext
        Loop
    End With
    If strObjs = "" Then Exit Function
    TestComponent = True
    SaveSetting "ZLSOFT", "注册信息", "本机部件", strObjs

End Function

Private Function MenuGranted(ByVal strMenuGroup As String) As ADODB.Recordset
    '-------------------------------------------------------------
    '功能：分析授权使用并安装的部件，进而产生授权使用的菜单集合
    '参数：注册码
    '-------------------------------------------------------------
    Dim ArrCommand
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCodes As String
    Dim strObjs As String
    Dim IntCount As Integer
    Dim strSystems As String
    Dim gstrMenuSys As String
    Dim BlnOnlySys As Boolean '只有报表系统
    Dim strSYS As String
    
    BlnOnlySys = (gstrSystems = "REPORT")
    If BlnOnlySys Then
        strSystems = " '0'"
        strSYS = strSystems
    Else
        strSystems = Replace(gstrSystems, "系统 in (", "")
        strSystems = Replace(strSystems, ")", "")
        strSYS = strSystems
        strSystems = Replace(strSystems, "','", ",")
    End If
    
    '--分析权限菜单--
    With rsTemp
        If Command() = "" Then
            gstrMenuSys = "缺省"
        Else
            ArrCommand = Split(Command(), " ")
            If UBound(ArrCommand) = 0 Then
                '仅仅包含菜单组别（如果含有/，表示是用户加密码的格式，如：zlhis/his）
                If InStr(1, ArrCommand(0), "/") = 0 Then
                    gstrMenuSys = ArrCommand(0)
                Else
                    gstrMenuSys = "缺省"
                End If
            Else
                '用户名、密码及菜单组别
                If UBound(ArrCommand) = 2 Then
                    gstrMenuSys = ArrCommand(2)
                Else
                    gstrMenuSys = "缺省"
                End If
            End If
        End If
        If strMenuGroup <> "" Then gstrMenuSys = strMenuGroup
        strObjs = GetSetting("ZLSOFT", "注册信息", "本机部件", "")
        If strObjs = "" Then strObjs = "'Zl9Common'"
        strObjs = Replace(strObjs, "','", ",")

        strSql = "SELECT 层次, Id AS 编号, Nvl(上级id, 0) AS 上级, 标题, Decode(Nvl(短标题,'空'),'空',标题,短标题) As 短标题, 快键, 说明, Nvl(模块, 0) AS 模块, Nvl(系统, 0) AS 系统, " & _
                 "        Nvl(图标, 0) AS 图标, nvl(部件,'0') as 部件, Decode(Upper(Rtrim(部件)), 'ZL9REPORT', 1, 0) AS 报表 " & _
                 " FROM TABLE(CAST(Zltools.f_Reg_Menu('" & gstrMenuSys & "', " & strSystems & ", " & strObjs & ") As " & _
                 " Zltools.t_Menu_Rowset)) " & _
                 " ORDER BY 层次, Id"

        If .State = adStateOpen Then .Close
        .Open strSql, gcnOracle, adOpenKeyset
    End With
    
    Set MenuGranted = rsTemp
    
End Function

Public Sub WriteToIni(ByVal Filename As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
''写INI文件
    Dim buff As String * 128
    buff = Trim(Value) + Chr(0)
    WritePrivateProfileString Section, Key, buff, Filename

End Sub

Public Function ReadFromIni(ByVal Filename As String, ByVal Section As String, ByVal Key As String) As String
''读INI文件
    Dim i As Long
    Dim buff As String * 128
    GetPrivateProfileString Section, Key, "", buff, 128, Filename
    i = InStr(buff, Chr(0))
    ReadFromIni = Trim(Left(buff, i - 1))
End Function
