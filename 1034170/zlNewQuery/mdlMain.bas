Attribute VB_Name = "mdlMain"
Option Explicit

Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_PENTIUM = 586
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064

Public Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Type IPINFO
     dwAddr As Long   ' IP address
    dwIndex As Long ' interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type

Private Const MAX_IP = 5   'To make a buffer... i dont think you have more than 5 ip on your pc..

Private Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO  'array of IP address entries
End Type

'Public gobjDemand As Object                '导航台
Public SplashObj As New frmSplash
Public gcnOracle As New ADODB.Connection    '公共数据库连接

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录

Public gstrUserFlag As String               '当前用户标志(两位表示)，第1位：是否DBA；第2位：系统所有者

Public gstrDbUser As String                 '当前数据库用户
Public gstrStation As String                '本工作站名称
Public gstrMenuSys As String                '系统菜单

'-----------------------------------------
'发行码、注册码、发行码解析串、注册码解析串
Public gstrRegCode As String
Public gstrPublish As String
Public gstrParseRegCode As String
Public gstrParsePublish As String
'-----------------------------------------

Public gstrSystems As String

'取硬盘大小
Private Const DRIVE_UNKNOWN = 0
Private Const DRIVE_ABSENT = 1
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

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

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long


Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long
'---------------------------------------------------------------
'-注册表 API 声明...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1

Private Const STRSPLIT As String = "♂♂"
Private Const REGCMD As String = "REGEDIT /E"
Private Const REGFILE As String = "C:\REGFILE.REG"
Private Const REGDATA As String = "C:\REGDATA.REG"
Private Const REGDIRECTORY As String = """HKEY_CURRENT_USER\SOFTWARE\VB AND VBA PROGRAM SETTINGS\ZLSOFT"""

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

Public glngOld As Long, glngFormW As Long, glngFormH As Long

'---------------------------------------------------------------
'   授权、菜单、试用版本
'---------------------------------------------------------------
Public Sub Main()
    Dim lngReturn As Long
    Dim StrUnitName As String
    Dim BlnShowFlash As Boolean
    Dim strCode As String
    Dim IntCount As Integer
    Dim StrStyle As String
    Dim rsMenu As ADODB.Recordset
    Dim StrHaveSys As String
    
    '为实现XP风格，在显示窗体前必须执行该函数
    Call InitCommonControls
    
    BlnShowFlash = False
    Load SplashObj
    '由注册表中获取用户注册相关信息,如果用户单位名称不为空,则显示闪现窗体
    StrUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "提示", "")
    If StrUnitName <> "" Then
        With SplashObj
            '有两处需要处理
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            Call ApplyOEM_Picture(.imgPic, "PictureB")
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
    InitCommon gcnOracle
    If RegCheck = False Then
        Unload SplashObj
        Exit Sub
    End If
    
    '如果发行码无效（为空或为"-"），则退出
    gstrParsePublish = zlRegInfo("产品简名")
    gstrParseRegCode = zlRegInfo("单位名称", , -1)
    
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
            .lbl技术支持商.Caption = zlRegInfo("技术支持商", , -1)
            .LblProductName = zlRegInfo("产品标题")
            
            strCode = zlRegInfo("产品开发商", , -1)
            .lbl开发商.Caption = ""
            For IntCount = 0 To UBound(Split(strCode, ";"))
                .lbl开发商.Caption = .lbl开发商.Caption & Split(strCode, ";")(IntCount) & vbCrLf
            Next
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            .Show
            BlnShowFlash = True
        End If
        DoEvents
    End With
    
    '将用户注册相关信息写入注册表,供下次启动时显示
    SaveSetting "ZLSOFT", "注册信息", "单位名称", gstrParseRegCode
    SaveSetting "ZLSOFT", "注册信息", "产品全称", zlRegInfo("产品标题")
    SaveSetting "ZLSOFT", "注册信息", "产品名称", zlRegInfo("产品简名")
    SaveSetting "ZLSOFT", "注册信息", "技术支持商", zlRegInfo("技术支持商", , -1)
    SaveSetting "ZLSOFT", "注册信息", "开发商", zlRegInfo("产品开发商", , -1)
    SaveSetting "ZLSOFT", "注册信息", "WEB支持商简名", zlRegInfo("支持商简名")
    SaveSetting "ZLSOFT", "注册信息", "WEB支持EMAIL", zlRegInfo("支持商MAIL")
    SaveSetting "ZLSOFT", "注册信息", "WEB支持URL", zlRegInfo("支持商URL")

    '-------------------------------------------------------------
    '检查本机安装部件
    '-------------------------------------------------------------
    If TestComponent = False Then
        MsgBox "您没有操作任何系统的权限，程序被迫退出！", vbInformation, gstrSysName
        Unload SplashObj
        Exit Sub
    End If
    
    '-------------------------------------------------------------
    '调用帐套选择窗体
    '-------------------------------------------------------------
    With FrmAccoutChoose
        gstrSystems = .Show_me
        If .BlnSelect = False Then
            Unload SplashObj
            Exit Sub
        End If
        StrHaveSys = gstrSystems
        
        If gstrSystems = "REPORT" Then
            gstrSystems = ""
        Else
            gstrSystems = " (系统 in (" & gstrSystems & ") Or 系统 Is NULL)"
        End If

        If gstrSystems = "" Then
            MsgBox "您没有操作任何系统的权限，程序被迫退出！", vbInformation, gstrSysName
            Unload SplashObj
            Exit Sub
        End If
        
    End With
    '-------------------------------------------------------------
    '分析菜单及部件
    '-------------------------------------------------------------
    If StrHaveSys = "" Then
        MsgBox "您没有操作任何系统的权限,程序被迫退出！", vbInformation, gstrSysName
        Unload SplashObj
        Exit Sub
    End If
    
    gstrSQL = "SELECT 系统 FROM zlPrograms WHERE 序号=1536 AND 系统 IN (" & StrHaveSys & ")"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMain")
    
    If gRs.BOF Then
        MsgBox "您没有操作任何系统的权限,程序被迫退出！", vbInformation, gstrSysName
        Unload SplashObj
        Exit Sub
    End If
    
    glngSys = gRs("系统").Value
    If InStr(1, GetPrivFunc(glngSys, 1536), "基本") <= 0 Then
        MsgBox "您没有操作任何系统的权限,程序被迫退出！", vbInformation, gstrSysName
        Unload SplashObj
        Exit Sub
    End If
    
    '-------------------------------------------------------------
    '创建同义词
    '-------------------------------------------------------------
    Call CreateSynonyms(glngSys, 1536)
    
    Call GetUserInfo
        
    Unload SplashObj
        
    If Not 是否允许使用本工作站 Then Exit Sub
        
    Call CodeMan(glngSys, 1536)
    
End Sub

Private Function CreateSynonyms(ByVal lngSys As Long, ByVal lngModul As Long)
    Dim strSQL As String
    '创建模块所需对象的同义词(如果已创建则不会再创建)
    On Error Resume Next
    strSQL = "Zl_Createsynonyms(" & lngSys & ")"
    zlDatabase.ExecuteProcedure strSQL, "创建同义词"
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
    gstrServerName = strServerName
    SetDbUser gstrDbUser
    
    gstrConnect = strServerName & ";" & strUserName & ";" & strUserPwd
    
    OraDataOpen = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
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
    If ErrCenter() = 1 Then Resume
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

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim RsTmp As New ADODB.Recordset
    
    Set RsTmp = zlDatabase.GetUserInfo
    
    UserInfo.用户名 = gstrDbUser
    UserInfo.姓名 = gstrDbUser
    If Not RsTmp.EOF Then
        UserInfo.ID = RsTmp!ID
        UserInfo.编号 = RsTmp!编号
        UserInfo.简码 = IIf(IsNull(RsTmp!简码), "", RsTmp!简码)
        UserInfo.姓名 = IIf(IsNull(RsTmp!姓名), "", RsTmp!姓名)
        UserInfo.用户名 = IIf(IsNull(RsTmp!用户名), "", RsTmp!用户名)
        UserInfo.部门ID = IIf(IsNull(RsTmp!部门ID), 0, RsTmp!部门ID)
        UserInfo.部门码 = IIf(IsNull(RsTmp!部门码), "", RsTmp!部门码)
        UserInfo.部门 = IIf(IsNull(RsTmp!部门名), "", RsTmp!部门名)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub CodeMan(lngSys As Long, ByVal lngModul As Long)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号

    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
       
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    
    glngSys = lngSys

    gstrUnitName = GetUnitName
'    gblnInsure = (UCase(GetSetting("ZLSOFT", "公共全局", "是否支持医保", "")) = UCase("Yes"))
'    gintInsure = Val(GetSetting("ZLSOFT", "公共全局", "医保类别", 0))
    
    gblnInsure = True
    Call gclsInsure.InitOracle(gcnOracle)
    '-------------------------------------------------
        
    frmMainQuery.Show
    
End Sub

Public Sub InitData()

End Sub

Public Function CloseChildWindows(ByVal frmMain As Object, ByVal FrmSon As Object) As Boolean
    '功能:关闭所有子窗口
    
    Dim FrmThis As Form
    
    On Error Resume Next

    CloseChildWindows = True
    
    For Each FrmThis In Forms
        If FrmThis.Caption <> frmMain.Caption And FrmThis.Caption <> FrmSon.Caption Then Unload FrmThis
    Next
    
    '关闭公共部件的窗体
    If CloseChildWindows Then CloseChildWindows = CloseWindows

End Function

Public Function TestComponent() As Boolean
    '如果没有任何部件可使用，则返回假
    TestComponent = False
    
    Dim strObjs As String, strCodes As String, strSQL As String
    Dim objComponent As Object
    Dim resComponent As New ADODB.Recordset
    
    '--由注册表获取授权部件--
    strObjs = GetSetting("ZLSOFT", "注册信息", "本机部件", "")
    If strObjs <> "" Then TestComponent = True: Exit Function
    '--分析授权安装部件--
    With resComponent
        strSQL = "Select Distinct Upper(g.部件) As 部件" & vbCrLf & _
                " From zlPrograms g, zlRegFunc r" & vbCrLf & _
                " Where g.序号 = r.序号 And Trunc(g.系统 / 100) = r.系统 And Upper(g.部件) <> 'ZL9REPORT'"
        
        If .State = adStateOpen Then .Close
        Set resComponent = zlDatabase.OpenSQLRecord(strSQL, "mdlMain")
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


Public Sub RunMudal(ByVal lngNO As Long)
    Select Case lngNO
    Case 1
        frmDefTable.Show , gfrmMain
    Case 2
        frmPicture.Show , gfrmMain
    Case 3
        frmDoctor.Show , gfrmMain
    Case 4
        frmAdvice.Show , gfrmMain
    Case 5
        frmDefQuery.Show , gfrmMain
    Case 6
        frmDefTree.Show , gfrmMain
    Case 7
'        If gblnInsure Then
'            If Not gclsInsure.InitInsure(gcnOracle) Then gblnInsure = False
'        End If
        
        Call gclsInsure.InitOracle(gcnOracle)
        
        frmMainQuery.Show , gfrmMain
    Case 8
        Call InitLocPar
        Call InitSysPar
        
        On Error Resume Next
        
        frmselectinfo.Show , gfrmMain
    Case 9
        frmLisPrinterSetup.Show , gfrmMain
    End Select
End Sub

Public Function 是否允许使用本工作站() As Boolean
    '编写人:朱玉宝 2003-03-09
    '功能:判断是否允许该工作站使用程序；如果需要替换本地参数，则执行替换操作；如果需要升级，则调用外壳程序，并关闭退出
    Dim objFileSys As New FileSystemObject
    Dim rsUse As New ADODB.Recordset
    Dim strSQL As String, strInfo As String, strPath As String, strExeName As String
    Dim blnAllow As Boolean, blnUpdate As Boolean, int连接数 As Integer
    Dim str升级程序 As String, Error As Long
    
    On Error Resume Next
    Err = 0
    blnAllow = False
    blnUpdate = False
    str升级程序 = "zlHisCrust.exe"
    是否允许使用本工作站 = False
    strExeName = GetSetting("ZLSOFT", "公共全局", "执行文件", "")
    
    '判断是否允许使用
    strInfo = AnalyseConfigure
    strSQL = "Select Nvl(禁止使用,0) as 允许,Nvl(升级标志,0) as 升级,Nvl(收集标志,0) as 收集,连接数 From zlClients Where 工作站='" & AnalyseComputer & "'"
    
    rsUse.CursorLocation = adUseClient
    Set rsUse = zlDatabase.OpenSQLRecord(strSQL, "mdlMain")
    With rsUse
        If .EOF Then
            '还没有该工作站的数据，上传（IP、机器名、CPU、内存、硬盘、操作系统）
            strSQL = " Insert into zlClients" & _
                     " (IP,工作站,CPU,内存,硬盘,操作系统,部门)" & _
                     " Values " & _
                     "('" & Split(strInfo, STRSPLIT)(0) & "','" & Split(strInfo, STRSPLIT)(1) & _
                     "','" & Split(strInfo, STRSPLIT)(2) & "','" & Split(strInfo, STRSPLIT)(3) & _
                     "','" & Split(strInfo, STRSPLIT)(4) & "','" & Split(strInfo, STRSPLIT)(5) & _
                     "','" & UserInfo.部门 & "')"
            gcnOracle.Execute strSQL
            是否允许使用本工作站 = True
            Exit Function
        Else
            blnAllow = IIf(IIf(IsNull(!允许), 0, !允许) = 0, True, False)
            int连接数 = IIf(IsNull(!连接数), 0, !连接数) '0-表示无限制
            blnUpdate = IIf(IIf(IsNull(!升级), 0, !升级) = 1, True, False)
            If Not blnUpdate Then blnUpdate = (IIf(IsNull(!收集), 0, !收集) = 1)
        End If
    End With
    If Not blnAllow Then
        MsgBox "该工作站已被管理员禁用！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '连接数检查限制
    If int连接数 > 0 Then
        strSQL = "Select SID From v$Session Where Upper(PROGRAM) Like 'ZLHIS%.EXE' And Status<>'KILLED' And MACHINE=(Select MACHINE From v$Session Where AUDSID=UserENV('SessionID'))"
        If rsUse.State = 1 Then rsUse.Close
        Set rsUse = zlDatabase.OpenSQLRecord(strSQL, "mdlMain")
        If rsUse.RecordCount > int连接数 Then
            MsgBox "当前工作站最多只允许 " & int连接数 & " 个登录连接，当前已经有 " & rsUse.RecordCount - 1 & " 个连接。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    On Error GoTo errHand
    '如果存在需要更新的本机参数，则更新本机注册表
    If Not RegRestoreByManager Then Exit Function
    
    '如果需要升级，则调用外壳程序
    If blnUpdate Then
        On Error Resume Next
        
        strPath = objFileSys.GetParentFolderName(App.Path)
        Error = Shell(strPath & "\" & str升级程序 & " " & gcnOracle.ConnectionString & "||0||" & strExeName, vbNormalFocus)
        '调用外壳程序
        If Error = 0 Then
            MsgBox "未找到客户端自动升级工具，请与软件提供商联系！", vbInformation, gstrSysName
            blnAllow = True
        Else
            'MsgBox "导航台将关闭，正为您启动客户端自动升级工具！", vbInformation, gstrSysName
            blnAllow = False
        End If
    End If
    
    是否允许使用本工作站 = blnAllow
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function AnalyseConfigure() As String
    '编写人:朱玉宝 2003-03-09
    '功能:分析出本机的配置（IP、机器名、CPU、内存、硬盘、操作系统）
    Dim strCPU As String           'CPU
    Dim strMemory As String        '内存
    Dim strOS As String            '操作系统
    Dim strComputerName As String  '计算机名
    Dim strHD As String            '硬盘
    Dim strIP As String            'IP地址
    Dim verinfo As OSVERSIONINFO
    Dim sysinfo As SYSTEM_INFO
    Dim memsts As MEMORYSTATUS
    Dim memory&
    
    strIP = AnalyseIP
    
    '获取计算机名
    strComputerName = AnalyseComputer
    
    '获取硬盘信息
    strHD = AnalyseHardDisk
    
    ' 获得操作系统信息
    strOS = GetVersionInfo
     
    ' 获得CPU类型
    GetSystemInfo sysinfo
    Select Case sysinfo.dwProcessorType
    Case PROCESSOR_INTEL_386
        strCPU = "Intel 386"
    Case PROCESSOR_INTEL_486
        strCPU = "Intel 486"
    Case PROCESSOR_INTEL_PENTIUM
        strCPU = "Intel Pentium"
    Case PROCESSOR_MIPS_R4000
        strCPU = "MIPS R4000"
    Case PROCESSOR_ALPHA_21064
        strCPU = "DEC Alpha 21064"
    Case Else
        strCPU = "(unknown)"
    End Select
     
    ' 获得剩余内存
    GlobalMemoryStatus memsts
    memory& = memsts.dwTotalPhys
    strMemory = Format$(memory& \ 1024 \ 1024, "###,###,###") + "M"
    'strMemory = "Total Physical Memory: "
    'strMemory = strMemory + Format$(memory& \ 1024, "###,###,###") + "K"
'    memory& = memsts.dwAvailPhys
'    strMemory = strMemory + "Available Physical Memory: "
'    strMemory = strMemory + Format$(memory& \ 1024, "###,###,###") + "K"
'    memory& = memsts.dwTotalVirtual
'    strMemory = strMemory + "Total Virtual Memory: "
'    strMemory = strMemory + Format$(memory& \ 1024, "###,###,###") + "K"
'    memory& = memsts.dwAvailVirtual
'    strMemory = strMemory + "Available Virtual Memory: "
'    strMemory = strMemory + Format$(memory& \ 1024, "###,###,###") + "K"
    
    AnalyseConfigure = strIP & STRSPLIT & strComputerName & STRSPLIT & strCPU & _
                       STRSPLIT & strMemory & STRSPLIT & strHD & STRSPLIT & strOS
End Function

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
End Function

Public Function RegRestoreByManager() As Boolean
    '最后修改:刘兴宏,因改为方案形式，因此更改了此过程
    '编写人:朱玉宝 2003-03-08
    '功能:从数据库中恢复该用户或整个注册表
    Dim strSection As String, strKey As String, strData As String
    Dim rsReg As New ADODB.Recordset
    Dim rsParaList As New ADODB.Recordset
    Dim strSQL As String, strComputerName As String
    Dim blnUser As Boolean '是否正在用户的升级
    Dim blnSharedUser As Boolean '是否存在不分站点的用户，如果有，则需要在Zlclientparaset中插入记录。恢复标志为2
    On Error GoTo errHand
    RegRestoreByManager = False
    
    strComputerName = AnalyseComputer

    strSQL = "Select 方案号,工作站,用户名" & _
            " From Zlclientparaset " & _
            " Where ((工作站 = '" & strComputerName & "' And 用户名 Is Null ) Or  " & _
            "         (工作站 Is Null And 用户名='" & UCase(gstrDbUser) & "') Or  " & _
            "         (工作站='" & strComputerName & "' And 用户名='" & UCase(gstrDbUser) & "')) " & _
            "               And Nvl(恢复标志, 0) = 1"
    
    zlDatabase.OpenRecordset rsReg, strSQL, "获取是否存在参数恢复"
    
    If rsReg.RecordCount = 0 Then
        '不存在自动恢复的参数
        RegRestoreByManager = True
        Exit Function
    End If
    
    rsReg.Filter = "用户名='" & UCase(gstrDbUser) & "'"
    blnUser = rsReg.RecordCount <> 0
    '检查该用户名下是否在本站点升过级的
    blnSharedUser = False
    strSQL = ""
    
    If blnUser Then
        strSQL = "" & _
            "   Select 用户名 From Zlclientparaset " & _
            "   where 方案号=" & Val(zlCommFun.Nvl(rsReg!方案号)) & _
            "         and 工作站='" & strComputerName & "'  And 用户名='" & UCase(gstrDbUser) & "'" & _
            "         And 恢复标志=2"
        zlDatabase.OpenRecordset rsParaList, strSQL, "检查该用户是否在本站点中升级过"
        If rsParaList.RecordCount <> 0 Then
            blnSharedUser = True
        Else
            strSQL = ",'私有全局','私有模块'"
        End If
    End If
    rsReg.Filter = "用户名=null and 工作站<> null "
    If rsReg.RecordCount <> 0 Then
        strSQL = " And 类别 in ('公共全局', '公共模块'" & strSQL & ")"
        rsReg.MoveFirst
    Else
        If strSQL = "" Then
            Exit Function
        Else
            strSQL = " And 类别 in (" & Mid(strSQL, 2) & ")"
        End If
    End If
    rsReg.Filter = 0
    strSQL = "" & _
        "   Select 方案号,序号,类别,目录,键名,键值,参数来源,参数说明 " & _
        "   From zlClientparaList " & _
        "   where 方案号=" & Val(zlCommFun.Nvl(rsReg!方案号)) & strSQL
    zlDatabase.OpenRecordset rsParaList, strSQL, "恢复参数"
    With rsParaList
        Do While Not .EOF
            Select Case zlCommFun.Nvl(!类别)
            Case "公共全局", "公共模块"
                Call SaveSetting("ZLSOFT", !类别 & IIf(IsNull(!目录), "", "\" & !目录), !键名, IIf(IsNull(!键值), "", !键值))
            Case "私有全局", "私有模块"
                Call SaveSetting("ZLSOFT", !类别 & "\" & UCase(gstrDbUser) & IIf(IsNull(!目录), "", "\" & !目录), zlCommFun.Nvl(!键名), zlCommFun.Nvl(!键值))
            End Select
            .MoveNext
        Loop
    End With
    
    
    '更新站点信息恢复
    'zl_zlClientParaSet_Restore(
    '    方案号_IN   IN zlClientParaSet.方案号%type,
    '    工作站_IN   IN zlClientParaSet.工作站%type,
    '    用户名_IN   IN zlClientParaSet.用户名%TYPE
    strSQL = "zl_zlClientParaSet_Restore("
    strSQL = strSQL & "" & Val(zlCommFun.Nvl(rsReg!方案号)) & ","
    strSQL = strSQL & "'" & UCase(strComputerName) & "',"
    strSQL = strSQL & "'" & UCase(gstrDbUser) & "')"
    zlDatabase.ExecuteProcedure strSQL, "保存恢复标志"
    RegRestoreByManager = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function AnalyseIP() As String
Dim ret As Long, Tel As Long
Dim bBytes() As Byte
Dim TempList() As String
Dim TempIP As String
Dim Tempi As Long
Dim Listing As MIB_IPADDRTABLE
Dim L3 As String


On Error GoTo END1
    GetIpAddrTable ByVal 0&, ret, True


    If ret <= 0 Then Exit Function
    ReDim bBytes(0 To ret - 1) As Byte
    ReDim TempList(0 To ret - 1) As String
    
    'retrieve the data
    GetIpAddrTable bBytes(0), ret, False
      
    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory Listing.dEntrys, bBytes(0), 4
    
    For Tel = 0 To Listing.dEntrys - 1
        'Copy whole structure to Listing..
        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
        TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
    Next Tel
    'Sort Out The IP For WAN
        TempIP = TempList(0)
        For Tempi = 0 To Listing.dEntrys - 1
            L3 = Left(TempList(Tempi), 3)
            If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
                TempIP = TempList(Tempi)
            End If
        Next Tempi
        AnalyseIP = TempIP 'Return The TempIP


Exit Function
END1:
AnalyseIP = ""
End Function

Public Function AnalyseHardDisk() As String
    '编写人:朱玉宝 2003-03-09
    '功能:获取硬盘总容量
    Dim lngSec As Long, lngByte As Long, lngFree As Long, lngClus As Long
    Dim strDrive As String, dblSum As Double
    
    strDrive = "C"
    Do Until strDrive > "Z"
        If GetDriveType(strDrive & ":\") = DRIVE_FIXED Then
            If GetDiskFreeSpace(strDrive & ":\", lngSec, lngByte, lngFree, lngClus) <> 0 Then
                dblSum = dblSum + lngSec * lngByte * CDbl(lngClus)
            End If
        End If
        
        strDrive = Chr(Asc(strDrive) + 1)
    Loop
    AnalyseHardDisk = Format(dblSum / 1024 / 1024 / 1024, "0.00") & "G"
End Function

Private Function GetVersionInfo() As String
    Dim myOS As OSVERSIONINFOEX
    Dim bExInfo As Boolean
    Dim sOS As String
    
    '如果是Windows2000或以下版本，则用新API再取一次
    myOS.dwOSVersionInfoSize = Len(myOS) 'should be 148/156
    'try win2000 version
    If GetVersionEx(myOS) = 0 Then
        'if fails
        myOS.dwOSVersionInfoSize = 148 'ignore reserved data
        If GetVersionEx(myOS) = 0 Then
            GetVersionInfo = "Windows (Unknown)"
            Exit Function
        End If
    Else
        bExInfo = True
    End If
    
    With myOS
        'is version 4
        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
            'nt platform
            Select Case .dwMajorVersion
            Case 3, 4
                sOS = "Windows NT"
            Case 5
                sOS = "Windows 2000"
            End Select
            If bExInfo Then
                'workstation/server?
                If .wProductType = VER_NT_SERVER Then
                    sOS = sOS & " Server"
                ElseIf .wProductType = VER_NT_DOMAIN_CONTROLLER Then
                    sOS = sOS & " Domain Controller"
                ElseIf .wProductType = VER_NT_WORKSTATION Then
                    sOS = sOS & IIf(.dwMajorVersion >= 5, " Professional", " WorkStation")
                End If
            End If
            
            'get version/build no
            'sOS = sOS & " Version " & .dwMajorVersion & "." & .dwMinorVersion & " " & TrimNull(.szCSDVersion) & " (Build " & .dwBuildNumber & ")"
            
        ElseIf .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
            'get minor version info
            If .dwMinorVersion = 0 Then
                sOS = "Windows 95"
            ElseIf .dwMinorVersion = 10 Then
                sOS = "Windows 98"
            ElseIf .dwMinorVersion = 90 Then
                sOS = "Windows Millenium"
            Else
                sOS = "Windows 9?"
            End If
            'get version/build no
            'sOS = sOS & "Version " & .dwMajorVersion & "." & .dwMinorVersion & " " & TrimNull(.szCSDVersion) & " (Build " & .dwBuildNumber & ")"
        End If
    End With
    GetVersionInfo = sOS
End Function

Public Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function


