Attribute VB_Name = "mdlMain"
Option Explicit

Public ZlBrowerDll As Object                '导航台
Public gcnOracle As New ADODB.Connection    '公共数据库连接
Public gobjRelogin As clsRelogin            '重新启动的类的对象实例
Public gobjWait As Object                   '展示非模态窗体后可以使程序不退出的对象
Public gstrCommand As String                '命令行内容


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

Public gstrSystems As String

Public gobjFile As New FileSystemObject

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'---------------------------------------------------------------
'-注册表 API 声明...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'切换到指定的输入法。
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long


Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long



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
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
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

Public Enum REGISTER
    注册信息
    私有模块
    私有全局
    公共模块
    公共全局
End Enum

'---------------------------------------------------------------
'启动时间，用以判断闪现屏幕的等待时间
'---------------------------------------------------------------
Public gdtStart As Long

'---------------------------------------------------------------
'   授权、菜单、试用版本
'---------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'功能:设置相关进程处理的API声明:2008-10-30 11:34:11:刘兴宏
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Type PROCESSENTRY32
      lSize             As Long
      lUsage            As Long
      lProcessId        As Long
      lDefaultHeapId    As Long
      lModuleId         As Long
      lThreads          As Long
      lParentProcessId  As Long
      lPriClassBase     As Long
      lFlags            As Long
      sExeFile          As String * 1024
End Type
Private Const PROCESS_TERMINATE = &H1
Public gcll_His_PId As Collection        '存储相关的进程信息:array(进程名称,PID,窗口个数),"K"+进程数

#Const SYS_TRYUSE = "正式" '正式/试用


Private Sub SetAppBusyState()
'当其他进程对象未创建完成时，替换在执行主进程功能时弹出的“部件被挂起”对话框
On Error Resume Next
    App.OleServerBusyMsgTitle = App.ProductName
    App.OleRequestPendingMsgTitle = App.ProductName
    
    App.OleServerBusyMsgText = "相关组件正在创建，请耐心等待。"
    App.OleRequestPendingMsgText = "相关组件正创建，请耐心等待。"
    
    App.OleServerBusyTimeout = 3000
    App.OleRequestPendingTimeout = 10000
Err.Clear
End Sub

Public Sub Main()
    Dim lngReturn As Long
    Dim StrUnitName As String
    Dim BlnShowFlash As Boolean
    Dim strCode As String, intCount As Integer, strStyle As String, strPath As String
    Dim strTitle As String                  '产品标题
    Dim strTag As String                    '旗舰版标志
    Dim rsMenu As ADODB.Recordset
    Dim objRIS As Object
    
    gstrCommand = CStr(Command())
    Set gobjRelogin = New clsRelogin
    gobjRelogin.MenuGroup = GetMenuGroup(gstrCommand)
    Call SetAppBusyState
     '为实现XP风格，在显示窗体前必须执行该函数
    Call InitCommonControls
    BlnShowFlash = False
    If InStr(gstrCommand, "=") <= 0 Then Load frmSplash
    '由注册表中获取用户注册相关信息,如果用户单位名称不为空,则显示闪现窗体
    StrUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "提示", "")
    If StrUnitName <> "" And StrUnitName <> "-" Then
        gdtStart = Timer
        With frmSplash
            '有两处需要处理
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            Call ApplyOEM_Picture(.imgPic, "PictureB")
            If InStr(gstrCommand, "=") <= 0 Then .Show
            .lblGrant = Replace(StrUnitName, ";", vbCrLf)
            StrUnitName = GetSetting("ZLSOFT", "注册信息", "开发商", "")
            If Trim(StrUnitName) = "" Then
                .Label3.Visible = False
                .lbl开发商.Visible = False
            Else
                .Label3.Visible = True
                .lbl开发商.Visible = True
                .lbl开发商.Caption = ""
                For intCount = 0 To UBound(Split(StrUnitName, ";"))
                    .lbl开发商.Caption = .lbl开发商.Caption & Split(StrUnitName, ";")(intCount) & vbCrLf
                Next
            End If
            .LblProductName = GetSetting("ZLSOFT", "注册信息", "产品全称", "")
            If Len(.LblProductName) > 10 Then
                .LblProductName.FontSize = 15.75 '三号
            Else
                .LblProductName.FontSize = 21.75 '二号
            End If
            .lbl技术支持商 = GetSetting("ZLSOFT", "注册信息", "技术支持商", "")
            .lbltag = GetSetting("ZLSOFT", "注册信息", "产品系列", "")
            
            If Trim$(.lbl技术支持商.Caption) = "" Then
                .Label1.Visible = False
                .lbl技术支持商.Visible = False
            Else
                .Label1.Visible = True
                .lbl技术支持商.Visible = True
            End If
        End With
        Do
            If (Timer - gdtStart) > 1 Then Exit Do
            DoEvents
        Loop
        
        BlnShowFlash = True
        DoEvents
    End If
    '问题:14365
    Call zlKillHISPID
    '用户注册
    If InStr(gstrCommand, ",") > 0 Or InStr(gstrCommand, "=") > 0 Or InStr(gstrCommand, "&") > 0 Then
        If Not frmUserLogin.Docmd(gstrCommand) Then
            If Not frmUserLogin.ShowMe Then
                Unload frmUserLogin
                Unload frmSplash
                Exit Sub
            End If
        End If
    Else
        If Not frmUserLogin.ShowMe Then
            Unload frmUserLogin
            Unload frmSplash
            Exit Sub
        End If
    End If

    If gcnOracle.State <> adStateOpen Then
        Unload frmUserLogin
        Unload frmSplash
        Exit Sub
    End If
    
    
    '写入本次启动程序的信息
    If IsDesinMode Then
        strPath = GetSetting("ZLSOFT", "公共全局", "程序路径", "")
        If strPath = "" Then
            strPath = "C:\Appsoft"
        Else
            strPath = Mid(strPath, 1, InStrRev(strPath, "\") - 1)
        End If
        gstrAviPath = strPath & "\附加文件"
    Else
        SaveSetting "ZLSOFT", "公共全局", "执行文件", App.EXEName & ".exe"
        SaveSetting "ZLSOFT", "公共全局", "程序路径", App.Path & "\" & App.EXEName & ".exe"
                
        gstrAviPath = App.Path & "\附加文件"
        SaveSetting "ZLSOFT", "注册信息", UCase("gstrAviPath"), gstrAviPath
    End If
    
    '2010-05-19 自动升级提前到授权检查前执行。
    If CheckAllowByTerminal = False Then
        Unload frmSplash
        Exit Sub
    End If
    '初始化公共部件
    InitCommon gcnOracle
    zl9ComLib.SetDbUser gobjRelogin.DBUser
    zl9ComLib.gstrNodeNo = gobjRelogin.NodeNo
    If RegCheck = False Then
        Unload frmSplash
        Exit Sub
    End If
    '版本检查
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
            MsgBox "授权性质不正确，程序被迫退出！", vbInformation, gstrSysName
            Unload frmSplash
            Exit Sub
    End Select
    
    gstrSysName = zlRegInfo("产品简名") & "软件"
    SaveSetting "ZLSOFT", "注册信息", "提示", gstrSysName
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrSysName"), gstrSysName
    gstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrVersion"), gstrVersion
    
    strTag = ""
    strTitle = zlRegInfo("产品标题")
    If strTitle <> "" Then
        If InStr(strTitle, "-") > 0 Then
            If Split(strTitle, "-")(1) = "Ultimate" Then
                strTag = "旗舰版"
            ElseIf Split(strTitle, "-")(1) = "Professional" Then
                strTag = "专业版"
            End If
        End If
    End If
    strTitle = Split(strTitle, "-")(0)
    With frmSplash
        If BlnShowFlash = False Then
            .lblGrant = Replace(zlRegInfo("单位名称", , -1), ";", vbCrLf)
            .lbl技术支持商.Caption = zlRegInfo("技术支持商", , -1)
            
            .LblProductName = strTitle
            .lbltag = strTag
            strCode = zlRegInfo("产品开发商", , -1)
            .lbl开发商.Caption = ""
            For intCount = 0 To UBound(Split(strCode, ";"))
                .lbl开发商.Caption = .lbl开发商.Caption & Split(strCode, ";")(intCount) & vbCrLf
            Next
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            If InStr(gstrCommand, "=") <= 0 Then .Show
            BlnShowFlash = True
        End If
        DoEvents
    End With
    '将用户注册相关信息写入注册表,供下次启动时显示
    SaveSetting "ZLSOFT", "注册信息", "单位名称", zlRegInfo("单位名称", , -1)
    SaveSetting "ZLSOFT", "注册信息", "产品全称", strTitle
    SaveSetting "ZLSOFT", "注册信息", "产品名称", zlRegInfo("产品简名")
    SaveSetting "ZLSOFT", "注册信息", "技术支持商", zlRegInfo("技术支持商", , -1)
    SaveSetting "ZLSOFT", "注册信息", "开发商", zlRegInfo("产品开发商", , -1)
    SaveSetting "ZLSOFT", "注册信息", "WEB支持商简名", zlRegInfo("支持商简名")
    SaveSetting "ZLSOFT", "注册信息", "WEB支持EMAIL", zlRegInfo("支持商MAIL")
    SaveSetting "ZLSOFT", "注册信息", "WEB支持URL", zlRegInfo("支持商URL")
    SaveSetting "ZLSOFT", "注册信息", "产品系列", strTag
    '检查本机安装部件
    If TestComponent = False Then
        MsgBox "您没有操作任何系统的权限，程序被迫退出！", vbInformation, gstrSysName
        Unload frmSplash
        Exit Sub
    End If
    '调用帐套选择窗体
    With FrmAccoutChoose
        gobjRelogin.Systems = .Show_me
        If .BlnSelect = False Then
            Unload frmSplash
            Exit Sub
        End If
        If gobjRelogin.Systems = "" Then
            MsgBox "您没有操作任何系统的权限，程序被迫退出！", vbInformation, gstrSysName
            Unload frmSplash
            Exit Sub
        End If
    End With
    Call GetUserInfo(IIf(gobjRelogin.Systems = "REPORT", 0, Replace(gobjRelogin.Systems, "'", "")))
    '登录信息读取
    gstrDeptName = gobjRelogin.DeptName
    gstrDbUser = gobjRelogin.DBUser
    gstrSystems = gobjRelogin.Systems
    '读取登录变量
    gstrUserFlag = IIf(gobjRelogin.IsSysOwner, "01", "00")
    gstrStation = ComputerName
    If gstrStation = "" Then
        gstrStation = "..."
    End If
    '分析菜单及部件
    Set rsMenu = MenuGranted(gobjRelogin.MenuGroup)
    If rsMenu.EOF Then
        MsgBox "您没有操作任何系统的权限,程序被迫退出！", vbInformation, gstrSysName
        Unload frmSplash
        Exit Sub
    End If
    '不用再创建公共同义词，公共的在安装和升级时创建，私有的在进入模块时调用
    '选择调用不同风格导航台
    On Error Resume Next
    Err = 0
    strStyle = zlDatabase.GetPara("导航台", , , "zlBrw")
    Set ZlBrowerDll = CreateObject(strStyle & ".Cls" & Mid(strStyle, 3))
    If Err <> 0 Then
        If strStyle = "ZLBRW" Then
            MsgBox "启动失败，主程序的相关文件丢失，请重新安装！", vbInformation, gstrSysName
            Unload frmSplash
            Exit Sub
        Else
            Err = 0
            Set ZlBrowerDll = CreateObject("ZLBRW.ClsBrw")
            If Err <> 0 Then
                MsgBox "启动失败，主程序的相关文件丢失，请重新安装！", vbInformation, gstrSysName
                Unload frmSplash
                Exit Sub
            End If
        End If
    End If
    On Error Resume Next
    Set objRIS = CreateObject("zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    If Not objRIS Is Nothing Then
        Call objRIS.SaveDBConnectInfo(gobjRelogin.InputUser, gobjRelogin.InputPwd, gobjRelogin.ServerName, gobjRelogin.IsTransPwd)
    End If
    '升级本地注册表参数值
    Call UpdateParameters
    Unload frmSplash
    '以下两句防止程序终止
    Set gobjWait = frmSelClient
    Load gobjWait
    Call ZlBrowerDll.SetEnvironment(gstrSysName, gstrVersion, gstrAviPath, _
                          gstrUserFlag, gstrDbUser, glngUserId, _
                          gstrUserCode, gstrUserName, gstrUserAbbr, _
                          glngDeptId, gstrDeptCode, gstrDeptName, _
                          gstrStation, gstrMenuSys, gstrCommand)
    Call ZlBrowerDll.InitBrower(gobjRelogin, gcnOracle, rsMenu)
End Sub

Public Function TestComponent() As Boolean
    '如果没有任何部件可使用，则返回假
    TestComponent = False
    
    Dim strObjs As String, strCodes As String, strSql As String
    Dim objComponent As Object
    Dim resComponent As New ADODB.Recordset
    
    On Error GoTo errH
    '--由注册表获取授权部件--
    strObjs = GetSetting("ZLSOFT", "注册信息", "本机部件", "")
    If strObjs <> "" Then
        If InStr(strObjs, "'ZL9REPORT'") = 0 Then
            If CreateComponent("ZL9REPORT.ClsREPORT") Then
                strObjs = strObjs & ",'ZL9REPORT'"
                SaveSetting "ZLSOFT", "注册信息", "本机部件", strObjs
            End If
        End If
        TestComponent = True
        Exit Function
    End If
    '--分析授权安装部件--
    strSql = "Select Distinct 部件 From (" & _
                " Select Upper(g.部件) As 部件" & _
                " From zlPrograms g, zlRegFunc r" & _
                " Where g.序号 = r.序号 And Trunc(g.系统 / 100) = r.系统" & _
                " Union " & _
                " Select Upper(部件) as 部件 From zlPrograms Where 序号 Between 10000 And 19999)"
    Set resComponent = zlDatabase.OpenSQLRecord(strSql, "")
    With resComponent
        Do While Not .EOF
            If CreateComponent(!部件 & ".Cls" & Mid(!部件, 4)) Then
                strObjs = strObjs & IIf(strObjs = "", "", ",") & "'" & !部件 & "'"
            End If
            .MoveNext
        Loop
    End With
    If strObjs = "" Then Exit Function
    TestComponent = True
    SaveSetting "ZLSOFT", "注册信息", "本机部件", strObjs
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CreateComponent(StrComponent) As Boolean
    Dim objComponent        As Object
On Error GoTo errH
    Set objComponent = CreateObject(StrComponent)
    CreateComponent = True
    Exit Function
errH:
    Err.Clear
    CreateComponent = False
    Exit Function
End Function

Public Function MenuGranted(ByVal strMenuGroup As String) As ADODB.Recordset
    '-------------------------------------------------------------
    '功能：分析授权使用并安装的部件，进而产生授权使用的菜单集合
    '参数：注册码
    '-------------------------------------------------------------
    Dim ArrCommand
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCodes As String
    Dim strObjs As String
    Dim intCount As Integer
    Dim strSystems As String
    Dim BlnOnlySys As Boolean '只有报表系统
    Dim strSYS As String
    
    On Error GoTo errH
    BlnOnlySys = (gstrSystems = "REPORT")
    If BlnOnlySys Then
        strSystems = "'0'"
        strSYS = "0"
    Else
        strSystems = Replace(gstrSystems, "','", ",")
        strSYS = Replace(gstrSystems, "'", "")
    End If
    
    If strMenuGroup <> "" Then gstrMenuSys = strMenuGroup
    strObjs = GetSetting("ZLSOFT", "注册信息", "本机部件", "")
    If strObjs = "" Then strObjs = "'Zl9Common'"
    strObjs = Replace(strObjs, "','", ",")
    If IsDesinMode Then
        strSql = "Select 层次, ID As 编号, Nvl(上级id, 0) As 上级, 标题, Decode(Nvl(短标题,'空'),'空',标题,短标题) as 短标题, 快键, 说明, Nvl(模块, 0) As 模块, Nvl(系统, 0) As 系统, " & _
                 "        Nvl(图标, 0) As 图标, 部件, Decode(Upper(RTrim(部件)), 'ZL9REPORT', 1, 0) As 报表 " & _
                 " From Table(Cast(ZLTOOLS.f_Reg_Menu([1], [2], [3]) As ZLTOOLS.t_Menu_Rowset)) " & _
                 " Union " & _
                 " Select A.层次, A.ID, Nvl(上级id, 0) As 上级, A.标题, Decode(Nvl(A.短标题,'空'),'空',A.标题,A.短标题) As 短标题, A.快键, A.说明, Nvl(A.模块, 0) As 模块, " & _
                 "        Nvl(A.系统, 0) As 系统, Nvl(图标, 0) As 图标, C.部件, Decode(C.部件, 'ZL9REPORT', 1, 0) As 报表 " & _
                 " From (Select Level As 层次, ID, 上级id, 标题, 短标题, 快键, 说明, Nvl(模块,0) 模块, 系统, 图标 " & _
                 "        From zlMenus " & _
                 "        Where 组别 = [1] And Nvl(系统, 0) IN(" & strSYS & ") " & _
                 "        Start With 上级id Is Null " & _
                 "        Connect By Prior ID = 上级id) A, " & _
                 "      (Select 系统, Nvl(模块,0) 模块 " & _
                 "        From zlMenus A " & _
                 "        Where 组别 = [1] And Nvl(系统, 0) IN (" & strSYS & ") " & _
                 "        Minus " & _
                 "        Select 系统 * 100, 序号 From Zlregfunc Where 系统 * 100 IN (" & strSYS & ")) B," & _
                 "      (select 系统, Upper(RTrim(部件)) as 部件,序号 From zlPrograms ) C " & _
                 " Where A.系统 = B.系统 And A.模块 = B.模块 And A.模块 = C.序号(+) and A.系统 = C.系统"

    Else
        strSql = "SELECT 层次, Id AS 编号, Nvl(上级id, 0) AS 上级, 标题, Decode(Nvl(短标题,'空'),'空',标题,短标题) As 短标题, 快键, 说明, Nvl(模块, 0) AS 模块, Nvl(系统, 0) AS 系统, " & _
                 "        Nvl(图标, 0) AS 图标, 部件, Decode(Upper(Rtrim(部件)), 'ZL9REPORT', 1, 0) AS 报表 " & _
                 " FROM TABLE(CAST(Zltools.f_Reg_Menu([1], [2], [3]) As " & _
                 " Zltools.t_Menu_Rowset)) "
    End If
    '实现报表按编号排序,模块号可能是zlReports.程序id,也可能是zlRPTGroups.程序id,优先zlReports
    '只获取报表发布到模块的报表
    strSql = "Select 层次, 编号, 上级, 标题, 短标题, 快键, 说明, 模块, 系统, 图标, 部件, 报表, 报表编号" & vbNewLine & _
                    "From (Select a.*, Decode(a.报表, 0, Null, Nvl(b.编号, c.编号)) 报表编号" & vbNewLine & _
                    "       From (" & strSql & ")  a," & vbNewLine & _
                    "            (Select b.系统, b.程序id, b.编号" & vbNewLine & _
                    "              From Zlprograms a, Zlreports b" & vbNewLine & _
                    "              Where Nvl(a.系统, 0) = Nvl(b.系统, 0) And a.序号 = Nvl(b.程序id, 0) And Upper(a.部件) = 'ZL9REPORT') b, Zlrptgroups c" & vbNewLine & _
                    "       Where a.系统 = b.系统(+) And a.模块 = b.程序id(+) And a.系统 = c.系统(+) And a.模块 = c.程序id(+))" & vbNewLine & _
                    "Order By 层次, 报表, 系统, 模块, 编号, 报表编号"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, gstrSysName, gstrMenuSys, Replace(strSystems, "'", ""), Replace(strObjs, "'", ""))

    Set MenuGranted = rsTemp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
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
    Dim strSql As String
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
            If InStr(strError, "ORA-00604") > 0 Then
                If InStr(strError, "ORA-20002") > 0 Then
                    strError = "当前用户不能使用该应用登录数据库，请联系管理员。"
                Else
                    strError = "当前用户被禁止登录数据库，请联系管理员。"
                End If
            End If
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
            ElseIf InStr(strError, "ORA-28001") > 0 Then
                MsgBox "密码已经过期。请联系管理员重置密码！", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo Errhand
    
    gstrDbUser = UCase(strUserName)
    SetDbUser gstrDbUser
    
    OraDataOpen = True
    Exit Function
    
Errhand:
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

Public Function ValEx(ByVal varInput As Variant) As Variant
'功能：由于Val只能以数字开头识别，ValEx以第一个数字进行识别
    Dim arrTmp As Variant, lngPos As Long
    If Val(varInput) = 0 Then
        varInput = varInput & ""
        If Trim(varInput) = "" Then ValEx = 0: Exit Function
        For lngPos = 1 To Len(varInput)
            If IsNumeric(Mid(varInput, lngPos, 1)) Then Exit For
        Next
        If lngPos = Len(varInput) + 1 Then
            ValEx = 0
        Else
            ValEx = Val(Mid(varInput, lngPos))
        End If
    Else
        ValEx = Val(varInput)
    End If
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

Public Function GetAllSubKey(ByVal KeyRoot As Long, KeyName As String) As Variant
'功能:获取某项的所有子项
'返回：=子项数组
    Dim lnghKey As Long, lngRet As Long, strName As String, lngIdx As Long
    Dim strSubKey As Variant
    strSubKey = Array()
    lngIdx = 0: strName = String(256, Chr(0))
    lngRet = RegOpenKey(KeyRoot, KeyName, lnghKey)
    If lngRet = 0 Then
        Do
            lngRet = RegEnumKey(lnghKey, lngIdx, strName, Len(strName))
            If lngRet = 0 Then
                ReDim Preserve strSubKey(UBound(strSubKey) + 1)
                strSubKey(UBound(strSubKey)) = Left(strName, InStr(strName, Chr(0)) - 1)
                lngIdx = lngIdx + 1
            End If
        Loop Until lngRet <> 0
    End If
    RegCloseKey lnghKey
    GetAllSubKey = strSubKey
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

Public Function ReadStartKey() As String
'功能：读取注册表中三个开始时间标志(之一有效即可)
    Dim strKey As String
    strKey = GetKeyValue(HKEY_CURRENT_USER, "SOFTWARE\VTCELUS6CS", "IXPHWP")  'FirstStart,1Start
    If strKey = "" Then strKey = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\EG5PZRELSML", "NXPHWP") 'SecondStart,2Start
    If strKey = "" Then strKey = GetKeyValue(HKEY_USERS, ".DEFAULT\SOFTWARE\S1NM9US6CS", "TXPHWP") 'ThirdStart,3Start
    If strKey <> "" Then ReadStartKey = CStr(CDate(strKey))
End Function

Public Function WriteStartKey() As Boolean
'功能:朝注册表中写三个开始时间标志
    Dim curDate As Date
    curDate = Format(Date, "yyyy-MM-dd")
    WriteStartKey = UpdateKey(HKEY_CURRENT_USER, "SOFTWARE\VTCELUS6CS", "IXPHWP", CCur(curDate)) 'FirstStart,1Start
    WriteStartKey = WriteStartKey And UpdateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\EG5PZRELSML", "NXPHWP", CCur(curDate)) 'SecondStart,2Start
    WriteStartKey = WriteStartKey And UpdateKey(HKEY_USERS, ".DEFAULT\SOFTWARE\S1NM9US6CS", "TXPHWP", CCur(curDate)) 'ThirdStart,3Start
End Function

Public Function ReadValidKey() As String
'功能：读取注册表中三个过期标志(之一有效即可)
    Dim strKey As String
    strKey = GetKeyValue(HKEY_CURRENT_USER, "SOFTWARE\PZ7Q64F9", "IRSUTR") 'OneValid,1Valid
    If strKey = "" Then strKey = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\SDDQ64F9", "NRSUTR") 'TwoValid,2Valid
    If strKey = "" Then strKey = GetKeyValue(HKEY_USERS, ".DEFAULT\SOFTWARE\S1CKGZHPNO", "TRSUTR") 'ThreeValid,3Valid
    If strKey <> "" Then ReadValidKey = strKey
End Function

Public Function WriteValidKey() As Boolean
    '功能:朝注册表中写三个过期标志
    WriteValidKey = UpdateKey(HKEY_CURRENT_USER, "SOFTWARE\PZ7Q64F9", "IRSUTR", "Q64F9") 'OneValid,1Valid
    WriteValidKey = WriteStartKey And UpdateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\SDDQ64F9", "NRSUTR", "Q64F9") 'TwoValid,2Valid
    WriteValidKey = WriteStartKey And UpdateKey(HKEY_USERS, ".DEFAULT\SOFTWARE\S1CKGZHPNO", "TRSUTR", "Q64F9") 'ThreeValid,3Valid
End Function

Public Function GetUserInfo(ByVal strSystems As String)
    Dim rsTmp As New ADODB.Recordset, rsUser As New ADODB.Recordset
    Dim strSql As String, i As Integer
    '读用户信息赋予公共，便于其他程序使用
    
    With rsTmp
        If .State = adStateOpen Then .Close
        strSql = "Select S.*" & _
                " From zlSystems S,(Select Distinct owner From All_Tables Where Table_Name='部门表') D" & _
                " Where Upper(S.所有者)=D.Owner And S.编号 In (" & strSystems & ") Order by S.编号"
        .Open strSql, gcnOracle, adOpenKeyset
        If Not .EOF Then
            '因为可能该用户具有多个系统的身份，所以循环取身份
            glngUserId = 0 '当前用户id
            gstrUserCode = "" '当前用户编码
            gstrUserName = "" '当前用户姓名
            gstrUserAbbr = "" '当前用户简码
            glngDeptId = 0 '当前用户部门id
            gstrDeptCode = "" '当前用户
            gstrDeptName = "" '当前用户
            
            For i = 1 To .RecordCount
                strSql = "Select R.*,D.编码 as 部门编码,D.名称 as 部门名称,P.编号,P.姓名,P.简码" & _
                        " From " & !所有者 & ".上机人员表 U," & !所有者 & ".人员表 P," & !所有者 & ".部门表 D," & !所有者 & ".部门人员 R" & _
                        " Where U.人员ID = P.ID And R.部门ID = D.ID And P.ID=R.人员ID and U.用户名=USER And (P.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.撤档时间 Is Null) and R.缺省=1"
                Set rsUser = New ADODB.Recordset
                rsUser.CursorLocation = adUseClient
                rsUser.Open strSql, gcnOracle, adOpenKeyset
                Set rsUser.ActiveConnection = Nothing
                If Not rsUser.EOF Then
                    glngUserId = rsUser!人员ID '当前用户id
                    gstrUserCode = rsUser!编号 '当前用户编码
                    gstrUserName = IIf(IsNull(rsUser!姓名), "", rsUser!姓名) '当前用户姓名
                    gstrUserAbbr = IIf(IsNull(rsUser!简码), "", rsUser!简码) '当前用户简码
                    glngDeptId = rsUser!部门ID '当前用户部门id
                    gstrDeptCode = rsUser!部门编码 '当前用户
                    gstrDeptName = rsUser!部门名称 '当前用户
                    Exit For
                End If
                DoEvents
                .MoveNext
            Next
        End If
        .Close
    End With
End Function

Private Function RunningInIDE() As Boolean
    '--检测是否源代码环境
    If App.EXEName = "prjMain" Then RunningInIDE = True
End Function

'**********************************************************************************************************************
'功能:以下处理相关进程的函数
'编制:刘兴洪
'日期:2008-10-30 11:38:58
Public Function zlKillHISPID() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:杀死所有HIS启动程序的移常进程(杀的条件是:所有ZLHIS+.exe的进程中无任何窗口)
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-30 11:06:16
    '-----------------------------------------------------------------------------------------------------------
    Dim lngProcess As Long, i As Long
    
    zlKillHISPID = False
    Err = 0: On Error GoTo Errhand:
    '第一步:需要处理相关的ZLHIS的相关进程
    Set gcll_His_PId = New Collection
    If zlHISPidToCollect(gcll_His_PId) = False Then zlKillHISPID = True: Exit Function  '如果存在相关的错误，就直接返回了
    If gcll_His_PId Is Nothing Then zlKillHISPID = True: Exit Function
    If gcll_His_PId.Count = 0 Then zlKillHISPID = True: Exit Function
    
    '第二步:需要处理相关ZLHIS的相关进程的相关窗口个数,这样才好判断出相关的进程是否存在异常,出现异常的，就得杀掉
    Call EnumWindows(AddressOf EnumWindowsProc, 0&)
    For i = 1 To gcll_His_PId.Count
        If Val(gcll_His_PId(i)(2)) <= 1 Then
            '肯定窗口数小于1或零,那么肯定有异常，需要杀死他
            If Val(gcll_His_PId(i)(1)) <> 0 Then
                '可能未成功，暂无处理此种情况
                Call TerminatePID(Val(gcll_His_PId(i)(1)))
            End If
        End If
    Next
    zlKillHISPID = True
Errhand:
End Function

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取所有窗口符合HIS的进程的窗口
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-30 10:26:02
    '-----------------------------------------------------------------------------------------------------------
    Dim strTittle As String, lngPID As Long, strName As String
    Dim lngCount As Long
    
    If GetParent(hwnd) = 0 Then
        '读取 hWnd 的视窗标题
        strTittle = String(80, 0)
        Call GetWindowText(hwnd, strTittle, 80)
        strTittle = Left(strTittle, InStr(strTittle, Chr(0)) - 1)
        If Trim(strTittle) <> "" Then
            Call GetWindowThreadProcessId(hwnd, lngPID)
            If IsWindowVisible(hwnd) Then
                Err = 0: On Error Resume Next
                strName = gcll_His_PId("K" & lngPID)(0)
                If Err = 0 Then
                    lngCount = Val(gcll_His_PId("K" & lngPID)(2)) + 1
                    gcll_His_PId.Remove "K" & lngPID
                    gcll_His_PId.Add Array(strName, lngPID, lngCount), "K" & lngPID
                End If
                Err.Clear: On Error GoTo 0
            End If
        End If
    End If
    EnumWindowsProc = True ' 表示继续列举 hWnd
    Exit Function
End Function

Private Function TerminatePID(ByVal lngPID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:结束指定的进程
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-30 11:06:16
    '-----------------------------------------------------------------------------------------------------------
    Dim lngProcess As Long
    TerminatePID = False
    
    Err = 0: On Error GoTo Errhand:
    lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPID)
    Call TerminateProcess(lngProcess, 1&)
    
    TerminatePID = True
Errhand:
End Function

Private Function zlHISPidToCollect(ByRef cll_His_Pid As Collection) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取ZLHIS的进程给相关的集合(gcll_HIS_Pid)
    '入参:
    '出参:cll_His_Pid-将符合HIS.exe的程序，装载该集合中
    '返回:
    '编制:刘兴洪
    '日期:2008-10-30 10:07:38
    '-----------------------------------------------------------------------------------------------------------
    Dim strEXEName  As String, lngSnapShot As Long, lngProcess As Long, lngCount  As Long
    Dim strCurExeName As String, lngCurPid As Long
    Dim uProcess   As PROCESSENTRY32
    Dim StrSessionID As String '当前会话ID
    Dim StrHISSessionID As String '其他ZLHIS进程会话ID
    Const TH32CS_SNAPPROCESS = &H2
    
    
    Err = 0: On Error GoTo Errhand:
    strCurExeName = "*" & UCase(App.EXEName) & "*"
    
    lngCurPid = GetCurrentProcessId '获取当前应用程序进程
    lngSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    
    StrSessionID = GetCurSessionID(lngCurPid)
    
    
    If lngSnapShot <> 0 Then
        uProcess.lSize = Len(uProcess)
        lngProcess = ProcessFirst(lngSnapShot, uProcess)
        lngCount = 0
        Do While lngProcess
            '不等于当前进程的才处理
            If lngCurPid <> uProcess.lProcessId Then
                strEXEName = UCase(Left(uProcess.sExeFile, InStr(1, uProcess.sExeFile, vbNullChar) - 1))
                If strEXEName Like strCurExeName Then '"ZLHIS+.EXE"
                    StrHISSessionID = GetCurSessionID(uProcess.lProcessId)
                    '如果当前zlhis+的进程会话ID与启动的会话ID相同,才进行关闭处理
                    If StrSessionID = StrHISSessionID Then
                        cll_His_Pid.Add Array(strEXEName, uProcess.lProcessId, 0), "K" & uProcess.lProcessId
                    End If
                End If
            End If
            lngProcess = ProcessNext(lngSnapShot, uProcess)
        Loop
        CloseHandle (lngSnapShot)
    End If
    zlHISPidToCollect = True
    Exit Function
Errhand:
End Function

Private Function GetCurSessionID(ByVal lngCurPid As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取当前进程的会话ID
    '入参:当前进程PID
    '出参:
    '返回:会话ID
    '编制:祝庆
    '日期:2012-06-06 10:15:00
    '-----------------------------------------------------------------------------------------------------------
    On Error Resume Next
    Dim WMI, objProcess, colProcessList As Object
    Set WMI = GetObject("WinMgmts:")
    Set colProcessList = WMI.InstancesOf("Win32_Process")
    For Each objProcess In colProcessList
        If objProcess.handle = lngCurPid Then
            GetCurSessionID = objProcess.SessionId
            Exit Function
        End If
    Next
    GetCurSessionID = "-1"
End Function

Public Function Is64bit() As Boolean
    '******************************************************************************************************************
    '功能：是否是64位系统
    '返回：
    '******************************************************************************************************************
    Dim handle As Long
    Dim bolFunc As Boolean
        
    bolFunc = False
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If handle > 0 Then
        IsWow64Process GetCurrentProcess(), bolFunc
    End If
    Is64bit = bolFunc
End Function

Public Function GetMenuGroup(ByVal strCommand As String) As String
    Dim ArrCommand As Variant
    '--分析权限菜单--
    If strCommand = "" Then
        GetMenuGroup = "缺省"
    Else
        ArrCommand = Split(gstrCommand, " ")
        If UBound(ArrCommand) = 0 Then
            '仅仅包含菜单组别（如果含有/，表示是用户加密码的格式，如：zlhis/his）
            If InStr(1, ArrCommand(0), "/") = 0 And InStr(ArrCommand(0), ",") = 0 Then
                GetMenuGroup = ArrCommand(0)
            Else
                GetMenuGroup = "缺省"
            End If
        Else
            '用户名、密码及菜单组别
            If UBound(ArrCommand) = 2 And InStr(ArrCommand(0), "=") <= 0 Then
                GetMenuGroup = ArrCommand(2)
            Else
                GetMenuGroup = "缺省"
            End If
        End If
    End If
End Function

