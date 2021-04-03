Attribute VB_Name = "mdlClient"
Option Explicit
 
Public gcnOracle As ADODB.Connection
Public gstrServerPath As String         '服务器目录
Public gstrSoftPath As String           '目标应用部件的所在目录
Public gstrWinSystemPath As String      '系统目录
Public gstrComputerIp As String         '本机的IP地址
Public gstrComputerName As String       '本机名
Public gstrAppPath  As String           '当前目录

Public gstrVisitUser As String       '访问的用户名
Public gstrVisitPassWord As String   '访问的密码
Public gstrVisitPort As String       '访问的端口
Public gstrConnectString As String
Public gBlnHisCrustCompare As Boolean '是否比较HISCUST外壳程序
Public gstrHisCommand   As String     'ZLHIS启动程序调用时传入的参数，回调时传回
Public gstrWinPath As String           'windows目录
Public gstrAppEXE   As String           '调用本外壳程序的文件
Public gblnPreUpgrade As Boolean        '是否为预升级
Public gblnOfficialUpgrade As Boolean   '是否为正式升级
Public gbln预升完成 As Boolean          '预升级是否完成

Public gstr收集类型 As String           '如:log;doc等
Public gbln收集 As Boolean
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

'显示当前运行的窗体的API声明
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GWL_STYLE = (-16)
Public Const WS_VISIBLE = &H10000000
Public Const WS_BORDER = &H800000
Private Const INTERNET_FLAG_PASSIVE = &H8000000     '启用被动传输

'关闭当前运行的窗体的API声明
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10

Const OFS_MAXPATHNAME = 128
Const OF_EXIST = &H4000


Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Public Declare Function RegComLib Lib "RegCom.dll" (ByVal fileName As String, ByRef result As String) As Long
Public Declare Function UnregComLib Lib "RegCom.dll" (ByVal fileName As String, ByRef result As String) As Long


Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type
Public Const INFINITE = -1&
Public Const SYNCHRONIZE = &H100000

Const NO_ERROR = 0
Const CONNECT_UPDATE_PROFILE = &H1
Const RESOURCETYPE_DISK = &H1
Const RESOURCETYPE_PRINT = &H2
Const RESOURCETYPE_ANY = &H0
Const RESOURCE_CONNECTED = &H1
Const RESOURCE_REMEMBERED = &H3
Const RESOURCE_GLOBALNET = &H2
Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Const RESOURCEDISPLAYTYPE_SERVER = &H2
Const RESOURCEDISPLAYTYPE_SHARE = &H3
Const RESOURCEUSAGE_CONNECTABLE = &H1
Const RESOURCEUSAGE_CONTAINER = &H2

Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias _
        "WNetAddConnection2A" _
        (lpNetResource As NETRESOURCE, _
        ByVal lpPassword As String, _
        ByVal lpUserName As String, _
        ByVal dwFlags As Long) As Long

Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias _
        "WNetCancelConnection2A" _
        (ByVal lpName As String, _
        ByVal dwFlags As Long, _
        ByVal fForce As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
                ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_BOTTOM = 1
Public Const HWND_TOP = 0
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const SWP_SHOWWINDOW = &H40
Public mobjFile As New FileSystemObject
Public mobjText As TextStream


 Private Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type
Public Declare Function apiOpenFile Lib "kernel32" Alias "OpenFile" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long


'取IP的API
Public Const MAX_ADAPTER_NAME_LENGTH         As Long = 256
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH  As Long = 128
Public Const MAX_ADAPTER_ADDRESS_LENGTH      As Long = 8
Public Const ERROR_SUCCESS  As Long = 0
Public Type IP_ADDRESS_STRING
    IpAddr(0 To 15)  As Byte
End Type
Public Type IP_MASK_STRING
    IpMask(0 To 15)  As Byte
End Type
Public Type IP_ADDR_STRING
    dwNext     As Long
    IpAddress  As IP_ADDRESS_STRING
    IpMask     As IP_MASK_STRING
    dwContext  As Long
End Type
Public Type IP_ADAPTER_INFO
  dwNext                As Long
  ComboIndex            As Long  '保留
  sAdapterName(0 To (MAX_ADAPTER_NAME_LENGTH + 3))        As Byte
  sDescription(0 To (MAX_ADAPTER_DESCRIPTION_LENGTH + 3)) As Byte
  dwAddressLength       As Long
  sIPAddress(0 To (MAX_ADAPTER_ADDRESS_LENGTH - 1))       As Byte
  dwIndex               As Long
  uType                 As Long
  uDhcpEnabled          As Long
  CurrentIpAddress      As Long
  IpAddressList         As IP_ADDR_STRING
  GatewayList           As IP_ADDR_STRING
  DhcpServer            As IP_ADDR_STRING
  bHaveWins             As Long
  PrimaryWinsServer     As IP_ADDR_STRING
  SecondaryWinsServer   As IP_ADDR_STRING
  LeaseObtained         As Long
  LeaseExpires          As Long
End Type
Public Declare Function GetAdaptersInfo Lib "iphlpapi.dll" _
    (pTcpTable As Any, pdwSize As Long) As Long


Const MAX_IP = 5   'To make a buffer... i dont think you have more than 5 ip on your pc..
Type IPINFO
     dwAddr As Long   ' IP address
    dwIndex As Long ' interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type
Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO  'array of IP address entries
End Type
Type IP_Array
    mBuffer As MIB_IPADDRTABLE
    BufferLen As Long
End Type
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
'converts a Long to a string

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Const PROCESS_TERMINATE = &H1

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As MODULEENTRY32) As Long

Private Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Byte
    modBaseSize As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * 1024
End Type

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
Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Private Const TH32CS_INHERIT = &H80000000

'ZQ
Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1

Public Type FILETIME
  dwLowDateTime     As Long
  dwHighDateTime    As Long
End Type

Public Type WIN32_FIND_DATA
  dwFileAttributes  As Long
  ftCreationTime    As FILETIME
  ftLastAccessTime  As FILETIME
  ftLastWriteTime   As FILETIME
  nFileSizeHigh     As Long
  nFileSizeLow      As Long
  dwReserved0       As Long
  dwReserved1       As Long
  cFileName         As String * MAX_PATH
  cAlternate        As String * 14
End Type

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public gintUpType  As Integer '升级方式 0共享 1FTP
Public gintGatherTYpe As Integer '收集方式 0共享 1FTP
Public gblnOk      As Boolean
Public gstrTempPath As String  '临时存放目录
Public gstrPerTempPath As String '预升级临时路径

Private Sub Main()
    Dim arrCommand As Variant
'Command =
'预升级：   ConnectionString||外壳调用(1-外壳调用的,0主程序调用的)||PreUpgrade||命令行参数
'定时升级： ConnectionString||外壳调用(1-外壳调用的,0主程序调用的)||OfficialUpgrade或带路径的ZLHIS+执行文件||命令行参数||USER=ZLHIS PASS=HIS(界面输入的密码)

'ConnectionString
'10.35.10以前：Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=30_TXYY_135"";Persist Security Info=True;User ID=zlhis;Password=HIS;Data Provider=MSDASQL
'10.35.10及以后：Provider=MSDataShape.1;Persist Security Info=False;User ID=ZLHIS;Data Provider=MSDASQL; 连接串中不含密码和服务器名了

    On Error Resume Next
    gstrConnectString = Command
    
    arrCommand = Split(gstrConnectString, "||")
    gBlnHisCrustCompare = True
    gstrConnectString = arrCommand(0)
    
    Select Case UBound(arrCommand)
    Case 0
    Case 1
        If arrCommand(1) = 1 Then
            gBlnHisCrustCompare = False
        End If
    Case Else
        gstrAppEXE = arrCommand(2)
        If arrCommand(1) = 1 Then
            gBlnHisCrustCompare = False
        End If
    End Select
    If UBound(arrCommand) = 4 Then
        gstrHisCommand = arrCommand(4)
    End If

    
    '是否预升级
    gblnOfficialUpgrade = False
    gblnPreUpgrade = (gstrAppEXE = "PreUpgrade")
    
    '正式升级
    If gstrAppEXE = "OfficialUpgrade" Or UCase(gstrAppEXE) = UCase("zlActMain.exe") Then
        gblnPreUpgrade = False
        gblnOfficialUpgrade = True
    End If
    
    If gblnPreUpgrade Then
        frmClientCopy.Hide
    Else
        frmClientCopy.Show
    End If
End Sub

Public Function IsUpgrade() As Boolean
'功能:确定是否需升级
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    IsUpgrade = False
    gbln收集 = False
    
    If gblnPreUpgrade Then
        strSQL = "Select 1 From zlClients Where  预升完成=1 and 工作站=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "检查预升", gstrComputerName)
        IsUpgrade = rsTmp.RecordCount = 0
        Exit Function
    Else
        strSQL = "Select 1 From zlClients Where  升级标志=1 and 工作站=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "检查是否升级", gstrComputerName)
        If rsTmp.RecordCount = 0 Then
            strSQL = "Select 1 From zlClients Where  收集标志=1 and 工作站=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, "检查是否收集", gstrComputerName)
            If rsTmp.RecordCount <> 0 Then
                gbln收集 = True
                IsUpgrade = True
            End If
        Else
            IsUpgrade = True
        End If
    End If
    
End Function

Public Function GetVersion(ByVal lngVer As Variant) As String
    '功能:将数字转换成版本号
    Dim strVer As String
    strVer = ""
    strVer = Int(lngVer / 10 ^ 8)
    If Len(lngVer) > 9 Then
        lngVer = Right(lngVer, 9) Mod (10 ^ 8)
    Else
        lngVer = lngVer Mod (10 ^ 8)
    End If
    
    lngVer = lngVer Mod 10 ^ 8
    strVer = strVer & "." & Int(lngVer / 10 ^ 4)
    lngVer = lngVer Mod 10 ^ 4
    strVer = strVer & "." & lngVer
    GetVersion = strVer
End Function

Public Function CompareFile(ByVal strSourceFile As String, ByVal strTargetFile As String, _
        Optional ByRef strSourceVer As String, Optional ByRef strSourceDate As String, _
        Optional ByRef strTargetVer As String, Optional ByRef strTargetDate As String _
        ) As Boolean

    '
    '功能:部件比较(比较版本号,修改时间)
    '入参数:
    '   strSourceFile:源文件
    '   strTargetFile:目标文件
    '出参数:
    '   strSourceVer:源版本号
    '   strSourceDate:源文件的最后修改日期
    '   strTargetVer:目标版本号
    '   strTargetDate:目标文件的最后修改日期
    '返回:如源文件比目标文件旧,则返回True,否则返回False

    Dim objFile As New FileSystemObject
    Dim strSource As String
    Dim strTarget As String
    
    CompareFile = False

    On Error Resume Next
    
    '比较文件版本号
    strSource = strSourceVer
    strTarget = GetCommpentVersion(strTargetFile)
    
    strSourceVer = strSource
    strTargetVer = strTarget
    If RtnVerNum(strTarget) < RtnVerNum(strSource) Then
        CompareFile = True
        
    End If
    
    
    '比较文件的最后修改时间
    
    strTarget = Format(FileDateTime(strTargetFile), "yyyy-MM-DD hh:mm:ss")
    If Err <> 0 Then
        strTarget = ""
        CompareFile = True
        Err = 0
    End If
    strSource = strSourceDate ' Format(FileDateTime(strSourceFile), "yyyy-MM-DD hh:mm:ss")
    
    
    strSourceDate = strSource
    strTargetDate = strTarget
    If strTarget < strSource Then
        CompareFile = True
    End If
End Function

Public Function GetCopyAndReg(ByVal strSourceFile As String, ByVal strTargetFile As String _
    , Optional ByRef strErrMsg As String = "升级正常！", Optional notReg As Boolean = False, Optional blnSysFile As Boolean = False, Optional notCopy As Boolean = False) As Boolean

    '功能:将源文拷给目标文件
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-01-20 13:46:50

    Dim objFile As New FileSystemObject, cllProcess As New Collection   '进程集array(进程,Exe文件名,模块进程)
    Dim strFile As String, intType As Integer '0-Exe可执行文件,1-DLL;OCX需要注册的文件,2-其他文件,如CHM,log等文件
    Dim i As Long, strMsgBox As String, iTask As Integer
    Dim pHandle As Long, ret As Integer
    Dim regMsg
    Dim strResult As String
    Dim lngResult As Long
    
    strFile = UCase(strSourceFile)
    If strFile Like "*.EXE" Then
        '判断是否为ACTIVEX EXE组件,是否需要注册
        If notReg Then
            intType = 3
        Else
            intType = 0
        End If
    ElseIf strFile Like "*.DLL" Or strFile Like "*.OCX" Then
        intType = 1
    Else
        intType = 2
    End If
    strFile = UCase(strTargetFile)

GoReExectue:
    On Error Resume Next
    
    '如果是文件类型=5,不强制升级,并且本地文件存在,直接判断为成功升级!
    If blnSysFile = True And notCopy = False And objFile.FileExists(strTargetFile) = True Then
        GoTo GoSuress
    End If
    
'    'Scrrun.dll 基础文件,特殊处理
'    If InStrRev(UCase(strSourceFile), UCase("\scrrun.dll")) > 0 Then
'        If objFile.FileExists(strTargetFile) Then
'            GoTo GoSuress
'        End If
'    End If
    
    '反注册
    If intType = 1 And notReg Then
'''        iTask = Shell("regsvr32 " & strTargetFile & "/u  /s", vbNormalFocus)
'''        pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
'''        ret = WaitForSingleObject(pHandle, INFINITE)
'''        ret = CloseHandle(pHandle)
'        Call UnRegister(strTargetFile)
        lngResult = UnregComLib(strTargetFile, strResult)
        If lngResult = 1 Or lngResult = 2 Then
            Call UnRegister(strTargetFile)
        End If
    End If
    
    'Active exe
    If intType = 3 And notReg Then
        lngResult = UnregComLib(strTargetFile, strResult)
        If lngResult = 1 Or lngResult = 2 Then
            Call UnRegServer(strTargetFile)
        End If
    End If
    
    
    '对于只读文件,需更改属性为普通属性
    If objFile.FileExists(strTargetFile) Then
        '文件存在,则改属性
        If FileSystem.GetAttr(strTargetFile) <> vbNormal Then
                FileSystem.SetAttr strTargetFile, vbNormal
        End If
    End If
    
    '检查本地目录是否存在
    Call CreatePath(strTargetFile)
    

    '拷贝文件
    objFile.CopyFile strSourceFile, strTargetFile, True
    If Err <> 0 Then
        '检查是否为系统文件
      
       strErrMsg = Err.Number & "-" & Err.Description

       '拒绝的权限
       If CheckSysFile(strTargetFile) Then
           '拒绝权限先改名
           
           If notCopy Then '强制覆盖
                Err.Clear
                
                On Error Resume Next
                Call Kill(strTargetFile & "_old")
                Name strTargetFile As strTargetFile & "_old"
                Call Kill(strTargetFile & "_old")
                
                '重新拷贝文件
                Err.Clear
                objFile.CopyFile strSourceFile, strTargetFile, True
                If Err <> 0 Then
                     If mobjFile.FileExists(strTargetFile) Then
                         iTask = Shell("regsvr32 " & strTargetFile & " /s", vbNormalFocus)
                         pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
                         ret = WaitForSingleObject(pHandle, INFINITE)
                         ret = CloseHandle(pHandle)
                     End If
                     GoTo GoSuress
                Else
                    strErrMsg = ""
                    GoTo Redo
                End If
           Else
                GoTo GoSuress
           End If
       End If
       
       '发生错误,肯定存在文件是只读或被独占打开或已执行
       If Err.Number <> 70 And Err.Number <> 70 - 2146828288 Then
            If MsgBox("注意：" & vbCrLf & _
                       "     文件“" & strTargetFile & "”，不能升级 ,原因如下：" & vbCrLf & Err.Number & "-" & Err.Description & vbCrLf & _
                       "『重试』表示手工已经解除相关错误，重新执行升级！" & vbCrLf & _
                       "『取消』表示取消本次升级！", vbQuestion + vbRetryCancel + vbDefaultButton1, "自动升级") = vbRetry Then
               '重新执行一次拷贝
               GoTo GoReExectue:
            Else
               Exit Function
            End If
       End If

        Select Case intType
        Case 0  ''0-Exe可执行文件
             Call zlGetFileProcess(strFile, cllProcess)
             strMsgBox = ""
             If gbln收集 Then
                 regMsg = MsgBox("注意：" & vbCrLf & _
                        "     文件“" & strTargetFile & "”可能没有写权限，不能往服务器进行收集 ！" & vbCrLf & _
                        "『终止』表示取消本文件收集！" & vbCrLf & _
                        "『重试』表示手工已经开启服务器目录权限，重新执行收集！" & vbCrLf & _
                        "『忽略』表示本次不进行收集！", vbQuestion + vbAbortRetryIgnore, "自动收集")
             Else
                 regMsg = MsgBox("注意：" & vbCrLf & _
                        "     文件“" & strTargetFile & "”正在执行，不能升级！" & vbCrLf & _
                        "『终止』表示取消本部件升级！" & vbCrLf & _
                        "『重试』表示终止被运行的程序，重新执行升级！" & vbCrLf & _
                        "『忽略』表示本次不进行升级！", vbQuestion + vbAbortRetryIgnore, "自动升级")    'vbAbortRetryIgnore
             End If
             
             If regMsg = 3 Then
                Exit Function
             ElseIf regMsg = 4 Then
                '先结束相关的进程
                For i = 0 To cllProcess.Count
                    Call TerminatePID(cllProcess(i)(0))
                Next
                '重新执行一次拷贝
                GoTo GoReExectue:
             Else
                strErrMsg = "忽略本部件升级"
                GoTo GoSuress
             End If
        Case 1  '1-DLL;OCX需要注册的文件
            strMsgBox = ""
            Call zlGetFileProcess(strFile, cllProcess)
            For i = 1 To cllProcess.Count
                If UCase(cllProcess(i)(1)) = UCase("ZLHISCRUST.EXE") Then
                    strErrMsg = "被自身独占"
                    GoTo GoSuress
                End If
                If i > 2 Then
                    strMsgBox = strMsgBox & Space(5) & cllProcess(i)(0) & "：" & cllProcess(i)(1) & vbCrLf & Space(5) & "...."
                    Exit For
                Else
                    strMsgBox = strMsgBox & Space(5) & cllProcess(i)(0) & "：" & cllProcess(i)(1) & vbCrLf
                End If
            Next
            '3终止
            '4重试
            '5忽略
            If gbln收集 Then
                regMsg = MsgBox("注意：" & vbCrLf & _
                        "     文件“" & strTargetFile & "”可能没有写权限，不能往服务器进行收集 ！" & vbCrLf & _
                        "『终止』表示取消本文件收集！" & vbCrLf & _
                        "『重试』表示手工已经开启服务器目录权限，重新执行收集！" & vbCrLf & _
                        "『忽略』表示本次不进行收集！", vbQuestion + vbAbortRetryIgnore, "自动收集")
            Else
                regMsg = MsgBox("注意：" & vbCrLf & _
                        "     文件“" & strTargetFile & "”正被如下程序引用，不能升级 ！" & vbCrLf & _
                        strMsgBox & vbCrLf & _
                        "『终止』表示取消本部件升级！" & vbCrLf & _
                        "『重试』表示终止被运行的程序，重新执行升级！" & vbCrLf & _
                        "『忽略』表示本次不进行升级！", vbQuestion + vbAbortRetryIgnore, "自动升级")    'vbAbortRetryIgnore
            End If
            
            If regMsg = 3 Then
                Exit Function
            ElseIf regMsg = 4 Then
               '先结束相关的进程
               For i = 0 To cllProcess.Count
                   Call TerminatePID(cllProcess(i)(0))
               Next
               '重新执行一次拷贝
               GoTo GoReExectue:
            Else
               strErrMsg = "忽略本部件升级"
               GoTo GoSuress
            End If
        Case Else   ',2-其他文件,如CHM,log等文件
            If gbln收集 Then
                regMsg = MsgBox("注意：" & vbCrLf & _
                        "     文件“" & strTargetFile & "”可能没有写权限，不能往服务器进行收集 ！" & vbCrLf & _
                        "『终止』表示取消本文件收集！" & vbCrLf & _
                        "『重试』表示手工已经开启服务器目录权限，重新执行收集！" & vbCrLf & _
                        "『忽略』表示本次不进行收集！", vbQuestion + vbAbortRetryIgnore, "自动收集")
            Else
                regMsg = MsgBox("注意：" & vbCrLf & _
                        "     文件“" & strTargetFile & "”正被其他文件独站打开，不能升级 ！" & vbCrLf & _
                        "『终止』表示取消本部件升级！" & vbCrLf & _
                        "『重试』表示手工已经解除独站运行的程序，重新执行升级！" & vbCrLf & _
                        "『忽略』表示本次不进行升级！", vbQuestion + vbAbortRetryIgnore, "自动升级")
            End If
            
            If regMsg = 3 Then
               Exit Function
            ElseIf regMsg = 4 Then
               '重新执行一次拷贝
               GoTo GoReExectue:
            Else
               strErrMsg = "忽略本部件升级"
               GoTo GoSuress
            End If
        End Select
    End If
Redo:
    If intType = 1 And notReg Then
        '1-DLL;OCX需要注册的文件,
'''        iTask = Shell("regsvr32 " & strTargetFile & " /s", vbNormalFocus)
'''        pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
'''        ret = WaitForSingleObject(pHandle, INFINITE)
'''        ret = CloseHandle(pHandle)
'        Call Register(strTargetFile)
        lngResult = RegComLib(strTargetFile, strResult)
        If lngResult = 1 Or lngResult = 2 Then
            Call Register(strTargetFile)
        End If
    End If
    
    
    If intType = 3 And notReg Then
        lngResult = RegComLib(strTargetFile, strResult)
        If lngResult = 1 Or lngResult = 2 Then
            Call RegServer(strTargetFile)
        End If
    End If
    
    strErrMsg = "正常升级!"
GoSuress:
    GetCopyAndReg = True
End Function

Public Function GetWinPath() As String
    '--功能:获取系统目录
    Dim Buffer As String
    Const MAX_PATH = 260
    Dim gstrWinPath As String
    Dim rtn As Long
    
    Buffer = Space(MAX_PATH)
    rtn = GetWindowsDirectory(Buffer, Len(Buffer))
    gstrWinPath = Left(Buffer, rtn)
    GetWinPath = gstrWinPath
End Function

Public Function GetWinSystemPath() As String
    
    Dim Buffer As String
    Dim strSystem As String
    Dim rtn As Long
    Const MAX_PATH = 260
    
    Buffer = Space(MAX_PATH)
    rtn = GetSystemDirectory(Buffer, Len(Buffer))
    strSystem = Left(Buffer, rtn)
    
    GetWinSystemPath = strSystem
End Function

Public Sub InintVar()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim str服务器号 As String
    
    On Error GoTo errH
    
    '初始化变量
    gstrWinSystemPath = GetWinSystemPath()
    gstrWinPath = GetWinPath()
    gstrSoftPath = GetSoftPath()
    
    If IsSourceCode Then
        gstrAppPath = "C:\APPSOFT"
    Else
        gstrAppPath = Replace(UCase(App.Path), UCase("\Apply"), "", 1)
    End If
    
    gstrComputerIp = AnalyseIP
    '获取工作站名
    gstrComputerName = GetMyCompterName
    
    If gblnPreUpgrade Then
        strSQL = "Select 升级服务器,FTP服务器 From zlClients Where 工作站=[1]"
    Else
        strSQL = "Select 升级服务器,FTP服务器 From zlClients Where  升级标志=1 and 工作站=[1]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "检查预升", gstrComputerName)
    
    With rsTmp
        If .RecordCount = 0 Then
            gbln收集 = Not gblnPreUpgrade
        Else
            gbln收集 = False
            If gintUpType = 0 Then
                If IsNull(rsTmp!升级服务器) Then
                    str服务器号 = ""
                Else
                    str服务器号 = rsTmp!升级服务器
                End If
            Else
                If IsNull(rsTmp!FTP服务器) Then
                    str服务器号 = ""
                Else
                    str服务器号 = rsTmp!FTP服务器
                End If
            End If
        End If
    End With
    
    
    If gbln收集 Then
        If gintGatherTYpe = 0 Then
            str服务器号 = "S"
            strSQL = "Select 项目,内容 From zlregInfo where 项目 in('收集目录S','访问用户S','访问密码S','收集类型')"
        Else
            str服务器号 = "F"
            strSQL = "Select 项目,内容 From zlregInfo where 项目 in('收集目录F','访问用户F','访问密码F','访问端口F','收集类型')"
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "初始化")
        With rsTmp
            Do While Not .EOF
                If !项目 = "收集目录" & str服务器号 Then
                    gstrServerPath = IIf(IsNull(!内容), "", !内容)
                End If
                If !项目 = "访问用户" & str服务器号 Then
                    gstrVisitUser = IIf(IsNull(!内容), "", !内容)
                End If
                If !项目 = "访问密码" & str服务器号 Then
                    gstrVisitPassWord = IIf(IsNull(!内容), "", !内容)
                End If
                If !项目 = "访问端口" & str服务器号 Then
                    gstrVisitPort = IIf(IsNull(!内容), "", !内容)
                End If
                If !项目 = "收集类型" Then
                    gstr收集类型 = IIf(IsNull(!内容), "", !内容)
                End If
                .MoveNext
            Loop
        End With
    Else
        '处理升级
        If gintUpType = 0 Then
            'Share处理方式
            If str服务器号 = "" Then
                strSQL = "Select 项目,内容 From zlregInfo where 项目 in('服务器目录','访问用户','访问密码')"
            Else
                strSQL = "Select 项目,内容 From zlregInfo where 项目 in('服务器目录" & str服务器号 & "','访问用户" & str服务器号 & "','访问密码" & str服务器号 & "')"
            End If
            Set rsTmp = OpenSQLRecord(strSQL, "初始化")
            With rsTmp
                Do While Not .EOF
                    If !项目 = "服务器目录" & str服务器号 Or !项目 = "收集目录" Then
                        gstrServerPath = IIf(IsNull(!内容), "", !内容)
                    End If
                    If !项目 = "访问用户" & str服务器号 Then
                        gstrVisitUser = IIf(IsNull(!内容), "", !内容)
                    End If
                    If !项目 = "访问密码" & str服务器号 Then
                        gstrVisitPassWord = IIf(IsNull(!内容), "", !内容)
                    End If
                    If !项目 = "收集类型" Then
                        gstr收集类型 = IIf(IsNull(!内容), "", !内容)
                    End If
                    .MoveNext
                Loop
            End With
            
            '20101206-zq 特殊处理服务器0 与 NULL 值为一样,解决服务器部件升级报错问题
            If gstrServerPath = "" And gbln收集 = False Then
                If str服务器号 = 0 Then
                     strSQL = "Select 项目,内容 From zlregInfo where 项目 in('服务器目录','访问用户','访问密码')"
                     Set rsTmp = OpenSQLRecord(strSQL, "检查服务器")

                     With rsTmp
                        Do While Not .EOF
                            If !项目 = "服务器目录" Or !项目 = "收集目录" Then
                                gstrServerPath = IIf(IsNull(!内容), "", !内容)
                            End If
                            If !项目 = "访问用户" Then
                                gstrVisitUser = IIf(IsNull(!内容), "", !内容)
                            End If
                            If !项目 = "访问密码" Then
                                gstrVisitPassWord = IIf(IsNull(!内容), "", !内容)
                            End If
                            If !项目 = "收集类型" Then
                                gstr收集类型 = IIf(IsNull(!内容), "", !内容)
                            End If
                            .MoveNext
                        Loop
                        
                     End With
                End If
            End If
        Else
            'FTP处理方式
            If str服务器号 = "" Then str服务器号 = "0"
            strSQL = "Select 项目,内容 From zlregInfo where 项目 in('FTP服务器" & str服务器号 & "','FTP用户" & str服务器号 & "','FTP密码" & str服务器号 & "','FTP端口" & str服务器号 & "')"
            Set rsTmp = OpenSQLRecord(strSQL, "检查FTP")
            
            With rsTmp
                Do While Not .EOF
                    If !项目 = "FTP服务器" & str服务器号 Or !项目 = "收集目录" Then
                        gstrServerPath = IIf(IsNull(!内容), "", !内容)
                    End If
                    If !项目 = "FTP用户" & str服务器号 Then
                        gstrVisitUser = IIf(IsNull(!内容), "", !内容)
                    End If
                    If !项目 = "FTP密码" & str服务器号 Then
                        gstrVisitPassWord = IIf(IsNull(!内容), "", !内容)
                    End If
                    If !项目 = "FTP端口" & str服务器号 Then
                        gstrVisitPort = IIf(IsNull(!内容), "", !内容)
                    End If
                    If !项目 = "收集类型" Then
                        gstr收集类型 = IIf(IsNull(!内容), "", !内容)
                    End If
                    .MoveNext
                Loop
            End With
        End If
    End If
    Exit Sub
    
errH:
    Call ErrCenter
End Sub

Public Sub AnalyseUserNameAndPassWord(strUser As String, strPassWord As String)
'--功  能:分析用户名或密码
'--入参数:
'--出参数:
'--返  回:

    Dim strArr
    Dim strTmp As String
    Dim strUserName As String
    Dim strPass As String
    Dim j As Long
    Dim i As Integer
    
    'Provider=MSDataShape.1;Extended Properties="Driver={Microsoft ODBC for Oracle};Server=ora2.domain";Persist Security Info=True;User ID=zlhis;Password=his;Data Provider=MSDASQL"
    strTmp = gstrConnectString
    strTmp = Replace(strTmp, """", "")
    strArr = Split(strTmp, ";")
    For i = 0 To UBound(strArr)
        j = InStr(1, UCase(strArr(i)), UCase("User ID"))
        If j <> 0 Then
            '取用户名
            strUserName = Mid(strArr(i), j + 8)
        End If
        j = InStr(1, UCase(strArr(i)), UCase("Password"))
        If j <> 0 Then
            '取用户名
            strPass = Mid(strArr(i), j + 9)
        End If
    Next
    strUser = strUserName
    strPassWord = strPass

End Sub

Public Function IsNetServer() As Boolean
    '--功能:检查服务器是否正常并连接
    Dim NetR As NETRESOURCE
    Dim objFile As New FileSystemObject
      
    '刘兴洪:可能存在windows资源管理器已经有访问的了
    '
    If objFile.FolderExists(gstrServerPath) Then
            IsNetServer = True: Exit Function
    End If
    
    If objFile.FolderExists(gstrServerPath) Then '存在此文件夹,肯定没有权限访问,则要删除连接
            Call zlNetCancelConnected '目前全部杀死,原因是不知道文件服务器名:如:IP和机器名访问
    End If
    
    
    With NetR
        .dwScope = RESOURCE_GLOBALNET
        .dwType = RESOURCETYPE_DISK
        .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
        .dwUsage = RESOURCEUSAGE_CONNECTABLE
        .lpLocalName = "" '映射的驱动器
        .lpRemoteName = gstrServerPath  '服务器路径
    End With
    
    On Error GoTo ErrHand:
    If WNetAddConnection2(NetR, gstrVisitPassWord, gstrVisitUser, CONNECT_UPDATE_PROFILE) = NO_ERROR Then
       IsNetServer = True
    Else
       IsNetServer = False
    End If
    Exit Function
ErrHand:
       IsNetServer = False
End Function

Public Function IsFtpServer() As Boolean
'功能:检查是否能正常连接FTP服务器
'                gstrServerPath
'                gstrVisitUser
'                gstrVisitPassWord
'                gstrVisitPort
        On Error GoTo errH
        If gstrServerPath = "" Or gstrVisitUser = "" Or gstrVisitPassWord = "" Or gstrVisitPort = "" Then
            IsFtpServer = False
            Exit Function
        End If
        
        glngINet = InternetOpen("FTP Control", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
        If glngINet <= 0 Then
            IsFtpServer = False
            Exit Function
        End If
    
        glngINetConn = InternetConnect(glngINet, gstrServerPath, gstrVisitPort, gstrVisitUser, gstrVisitPassWord, INTERNET_SERVICE_FTP, 0, 0)
        
        If glngINetConn Then
            IsFtpServer = True
        Else
            IsFtpServer = False
        End If
        Exit Function
errH:
    If Err Then
        IsFtpServer = False
    End If
End Function

Public Function CancelFtpServer() As Boolean
    On Error Resume Next
    If glngINetConn <> 0 Then
       InternetCloseHandle glngINetConn
    End If

    If glngINet <> 0 Then
       InternetCloseHandle glngINet
    End If
    CancelFtpServer = True
End Function

Public Function CancelNetServer(Optional strName As String) As Boolean
    '断开服务器连接
    Dim lngReturn As Long
    
    On Error Resume Next
    lngReturn = WNetCancelConnection2(IIf(strName = "", gstrServerPath, strName), CONNECT_UPDATE_PROFILE, True)
    If lngReturn = 0 Then
        CancelNetServer = True
    Else
        CancelNetServer = False
    End If
End Function


Private Function GetSoftPath() As String
    
    '功能:获取应用部件的所在目录
    
    Dim strFile As String
    Dim strPath As String
    Dim objFile As New FileSystemObject
    Dim blnRutn As Boolean
    Dim DirPath  As String
    Dim ll As FileListBox
    
    On Error GoTo ErrHand:
    '---确定当前系统目录
    blnRutn = False
    DirPath = Trim$(App.Path)
    GetSoftPath = DirPath
    
    If Right$(DirPath, 1) <> "\" Then
      DirPath = DirPath & "\"
    End If
    
    strPath = Dir$(DirPath & "*", vbArchive Or vbHidden Or vbSystem Or vbDirectory)
    Do
        
        DoEvents
        If strPath = "" Then
            Exit Do
        Else
            If blnRutn = False Then
                If (GetAttr(DirPath & strPath) And vbDirectory) = vbDirectory Then
                    If Left$(strPath, 1) <> "." And Left$(strPath, 2) <> ".." Then
                        With frmErrAsk.File1
                            .Path = DirPath & strPath
                            .fileName = "*.dll"
                            If .ListCount > 0 Then
                                GetSoftPath = DirPath & strPath
                                Exit Do
                            End If
                        End With
                End If
                End If
            End If
        End If
        strPath = Dir$
    Loop
ErrHand:
    
End Function

Public Function FindFile(ByVal strFileName As String) As Boolean
    '--功能:查找指定的文件是否存在
    '--返回: 如果存在此文件为True,否则为Flase
    Dim typOfStruct As OFSTRUCT
    
    On Error Resume Next
    FindFile = False
    If Len(strFileName) > 0 Then
        apiOpenFile strFileName, typOfStruct, OF_EXIST
        FindFile = typOfStruct.nErrCode <> 2
    End If
End Function
Public Function ExcuteHisProgram(ByVal strHisExeFile As String) As Long
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    ExcuteHisProgram = ShellExecute(Scr_hDC, "Open", strHisExeFile, "", "C:\", 1)
End Function

Public Function AnalyseIP() As String
    '功能:分析出本机的IP地址
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
Public Function TrimNull(item As String)
    Dim pos As Integer
    pos = InStr(item, Chr$(0))
    If pos Then
          TrimNull = Left$(item, pos - 1)
    Else: TrimNull = item
    End If
End Function
Public Function RtnVerNum(ByVal strVer As String) As Long

    '--功能:返回数字版本

    Dim strArr
    
    If strVer <> "" Then
        strArr = Split(strVer, ".")
        RtnVerNum = strArr(0) * 10 ^ 8 + strArr(1) * 10 ^ 4 + strArr(2)
    Else
        RtnVerNum = 0
    End If
End Function

Public Function isHisCurstUpGrade() As Boolean

    '--功能:确定是否升级外壳,并将其外壳拷入apply目录下

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strSourceVer As String
    Dim strSourceDate As String
    Dim strSourceMD5  As String
    Dim strTargetVer As String
    Dim strTargetDate As String
    Dim strTargetMD5  As String
    Dim strErrMsg As String
    Dim strTargetFile As String
    Dim strSourceFile  As String
    
    isHisCurstUpGrade = False
    
    
    On Error GoTo ErrHand:
    strSQL = "Select 修改日期,版本号,MD5 From zlfilesupgrade where upper(文件名)='ZLHISCRUST.EXE'"
    Set rsTmp = OpenSQLRecord(strSQL, "检查自升级")
    
    With rsTmp
        If .EOF() Then Exit Function
        
        strSourceFile = gstrAppPath & "\ZLHISCRUST.EXE"
        If gintUpType = 0 Then
            strTargetFile = gstrServerPath & "\ZLHISCRUST.EXE"
        Else
            strTargetFile = "ZLHISCRUST.EXE"
        End If
        
        If FindFile(strSourceFile) Then
            strSourceVer = GetCommpentVersion(strSourceFile)
            strSourceDate = Format(FileDateTime(strSourceFile), "yyyy-MM-DD hh:mm:ss")
'            strSourceFile = "D:\zlHisCrust.exe" 测试
            strSourceMD5 = HashFile(strSourceFile, 2 ^ 27)
        Else
            strSourceVer = "0.0.0"
            strSourceDate = ""
            strSourceMD5 = ""
        End If
        
        strTargetMD5 = NVL(!MD5, "")
         
         '比较MD5
         If UCase(strSourceMD5) <> UCase(strTargetMD5) Then
            strTargetVer = GetVersion(IIf(IsNull(!版本号), 0, !版本号))
            strTargetDate = Format(!修改日期, "yyyy-MM-DD HH:mm:ss")
            
            '判断目录是否存在,不存在自动创建.20101206--ZQ
            If mobjFile.FolderExists(gstrAppPath & "\Apply\") = False Then
               Call mobjFile.CreateFolder(gstrAppPath & "\Apply\")
            End If
            
            If gintUpType = 0 Then
                If GetCopyAndReg(strTargetFile, gstrAppPath & "\Apply\zlHisCrust.exe", strErrMsg) Then
                    isHisCurstUpGrade = True
                    WriteTxtLog strSourceFile & "(版本:" & strSourceVer & "   修改日期:" & strSourceDate & ")    ====>    " & _
                                    strTargetFile & "(版本:" & strTargetVer & "   修改日期:" & strTargetDate & ")        升级信息:" & strErrMsg
                    Exit Function
                End If
            Else
                If FtpDownFile(strTargetFile, gstrAppPath & "\Apply\zlHisCrust.exe") Then
                    isHisCurstUpGrade = True
                    WriteTxtLog strSourceFile & "(版本:" & strSourceVer & "   修改日期:" & strSourceDate & ")    ====>    " & _
                                    strTargetFile & "(版本:" & strTargetVer & "   修改日期:" & strTargetDate & ")        升级信息:" & strErrMsg
                    Exit Function
                End If
            End If
        End If
        
    End With
    Exit Function
ErrHand:
End Function

Public Function isRegComUpGrade(ByVal strRegName As String) As Boolean

    '--功能:确定是否升级外壳,并将其外壳拷入apply目录下

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strSourceVer As String
    Dim strSourceDate As String
    Dim strSourceMD5  As String
    Dim strTargetVer As String
    Dim strTargetDate As String
    Dim strTargetMD5  As String
    Dim strErrMsg As String
    Dim strTargetFile As String
    Dim strSourceFile  As String
    
    isRegComUpGrade = False
    
    
    On Error GoTo ErrHand:
    strSQL = "Select 修改日期,版本号,MD5 From zlfilesupgrade where upper(文件名)='" & strRegName & "'"
    Set rsTmp = OpenSQLRecord(strSQL, "检查自升级")
    
    With rsTmp
        If .EOF() Then Exit Function
        
        strSourceFile = gstrWinSystemPath & "\" & strRegName
        If gintUpType = 0 Then
            strTargetFile = gstrServerPath & "\" & strRegName
        Else
            strTargetFile = strRegName
        End If
        
        If FindFile(strSourceFile) Then
            strSourceVer = GetCommpentVersion(strSourceFile)
            strSourceDate = Format(FileDateTime(strSourceFile), "yyyy-MM-DD hh:mm:ss")
'            strSourceFile = "D:\zlHisCrust.exe" 测试
            strSourceMD5 = HashFile(strSourceFile, 2 ^ 27)
        Else
            strSourceVer = "0.0.0"
            strSourceDate = ""
            strSourceMD5 = ""
        End If
        
        strTargetMD5 = NVL(!MD5, "")
         
         '比较MD5
         If UCase(strSourceMD5) <> UCase(strTargetMD5) Then
            strTargetVer = GetVersion(IIf(IsNull(!版本号), 0, !版本号))
            strTargetDate = Format(!修改日期, "yyyy-MM-DD HH:mm:ss")
            
            '判断目录是否存在,不存在自动创建.20101206--ZQ
            If mobjFile.FolderExists(gstrAppPath & "\PUBLIC\") = False Then
               Call mobjFile.CreateFolder(gstrAppPath & "\PUBLIC\")
            End If
            
            If gintUpType = 0 Then
                If GetCopyAndReg(strTargetFile, gstrWinSystemPath & "\" & strRegName, strErrMsg) Then
                    isRegComUpGrade = True
                    WriteTxtLog strSourceFile & "(版本:" & strSourceVer & "   修改日期:" & strSourceDate & ")    ====>    " & _
                                    strTargetFile & "(版本:" & strTargetVer & "   修改日期:" & strTargetDate & ")        升级信息:" & strErrMsg
                    Exit Function
                End If
            Else
                If FtpDownFile(strTargetFile, gstrWinSystemPath & "\" & strRegName) Then
                    isRegComUpGrade = True
                    WriteTxtLog strSourceFile & "(版本:" & strSourceVer & "   修改日期:" & strSourceDate & ")    ====>    " & _
                                    strTargetFile & "(版本:" & strTargetVer & "   修改日期:" & strTargetDate & ")        升级信息:" & strErrMsg
                    Exit Function
                End If
            End If
        End If
        
    End With
    Exit Function
ErrHand:
    MsgBox "拷贝注册工具文件出错：" & vbCrLf & Err.Description, vbExclamation, "提示"
End Function

Public Sub isMD5UpGrade()

    '--功能:备用MD5DLL检查

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strErrMsg As String
    Dim strTargetFile As String
    Dim strSourceFile  As String
    
    
    On Error GoTo ErrHand:
    strSQL = "Select 修改日期,版本号,MD5 From zlfilesupgrade where upper(文件名)='AAMD532.DLL'"
    Set rsTmp = OpenSQLRecord(strSQL, "检查AAMD532")
    
    With rsTmp
        If .EOF() Then Exit Sub
        
        strSourceFile = gstrWinSystemPath & "\AAMD532.DLL"
        If gintUpType = 0 Then
            strTargetFile = gstrServerPath & "\AAMD532.DLL"
        Else
            strTargetFile = "AAMD532.DLL"
        End If
        
        If FindFile(strSourceFile) Then
            Exit Sub
        End If
                
        If gintUpType = 0 Then
            If GetCopyAndReg(strTargetFile, strSourceFile, strErrMsg) Then
                WriteTxtLog strSourceFile & " ====> " & strTargetFile & "升级信息:" & strErrMsg
                Exit Sub
            End If
        Else
            If FtpDownFile(strTargetFile, strSourceFile) Then
                WriteTxtLog strSourceFile & " ====> " & strTargetFile & "升级信息:" & strErrMsg
                Exit Sub
            End If
        End If
    End With
    
    Exit Sub
ErrHand:
    
End Sub

Public Sub is7zUpGrade()

    '--功能:处理7Z的压缩解压缩文件首先升级

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strSourceVer As String
    Dim strSourceDate As String
    Dim strSourceMD5  As String
    Dim strTargetVer As String
    Dim strTargetDate As String
    Dim strTargetMD5  As String
    Dim strErrMsg As String
    Dim strTargetFile As String
    Dim strSourceFile  As String
    
    
    On Error GoTo ErrHand:
    strSQL = "Select 文件名,修改日期,版本号,MD5 From zlfilesupgrade where upper(文件名)='7Z.EXE' or upper(文件名)='7Z.DLL'"
    Set rsTmp = OpenSQLRecord(strSQL, "检查7Z")
    
    With rsTmp
        Do Until rsTmp.EOF
            strSourceFile = gstrWinSystemPath & "\" & NVL(!文件名, "")
            If gintUpType = 0 Then
                strTargetFile = gstrServerPath & "\" & NVL(!文件名, "")
            Else
                strTargetFile = NVL(!文件名, "")
            End If
            
            If FindFile(strSourceFile) Then
                strSourceVer = GetCommpentVersion(strSourceFile)
                strSourceDate = Format(FileDateTime(strSourceFile), "yyyy-MM-DD hh:mm:ss")
                strSourceMD5 = HashFile(strSourceFile, 2 ^ 27)
            Else
                strSourceVer = "0.0.0"
                strSourceDate = ""
                strSourceMD5 = ""
            End If
            strTargetMD5 = NVL(!MD5, "")
             
             '比较MD5
             If UCase(strSourceMD5) <> UCase(strTargetMD5) Then
                strTargetVer = GetVersion(IIf(IsNull(!版本号), 0, !版本号))
                strTargetDate = Format(!修改日期, "yyyy-MM-DD HH:mm:ss")
                
                If gintUpType = 0 Then
                    If GetCopyAndReg(strTargetFile, strSourceFile, strErrMsg) Then
    
                        WriteTxtLog strSourceFile & "(版本:" & strSourceVer & "   修改日期:" & strSourceDate & ")    ====>    " & _
                                        strTargetFile & "(版本:" & strTargetVer & "   修改日期:" & strTargetDate & ")        升级信息:" & strErrMsg
                    End If
                Else
                    If FtpDownFile(strTargetFile, strSourceFile) Then
                        WriteTxtLog strSourceFile & "(版本:" & strSourceVer & "   修改日期:" & strSourceDate & ")    ====>    " & _
                                        strTargetFile & "(版本:" & strTargetVer & "   修改日期:" & strTargetDate & ")        升级信息:" & strErrMsg
                    End If
                End If
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHand:

End Sub


Public Sub isRunasUpGrade()

    '--功能:zlRunas检查

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strErrMsg As String
    Dim strTargetFile As String
    Dim strSourceFile  As String
    
    
    On Error GoTo ErrHand:
    strSQL = "Select 修改日期,版本号,MD5 From zlfilesupgrade where upper(文件名)='ZLRUNAS.EXE'"
    Set rsTmp = OpenSQLRecord(strSQL, "检查ZLRUNAS")
    
    With rsTmp
        If .EOF() Then Exit Sub
        
        strSourceFile = App.Path & "\ZLRUNAS.EXE"
        If gintUpType = 0 Then
            strTargetFile = gstrServerPath & "\ZLRUNAS.EXE"
        Else
            strTargetFile = "ZLRUNAS.EXE"
        End If
        
        If FindFile(strSourceFile) Then
            Exit Sub
        End If
                
        If gintUpType = 0 Then
            If GetCopyAndReg(strTargetFile, strSourceFile, strErrMsg) Then
                WriteTxtLog strSourceFile & " ====> " & strTargetFile & "升级信息:" & strErrMsg
                Exit Sub
            End If
        Else
            If FtpDownFile(strTargetFile, strSourceFile) Then
                WriteTxtLog strSourceFile & " ====> " & strTargetFile & "升级信息:" & strErrMsg
                Exit Sub
            End If
        End If
    End With
    
    Exit Sub
ErrHand:
    
End Sub



'converts a Long to a string
Public Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function
Public Function GetMyCompterName() As String
    '功能:获取计算机名
    '获取计算机名
    Dim strComputerName As String * 256
    
    On Error Resume Next
    
    Call GetComputerName(strComputerName, 255)
    GetMyCompterName = Trim(Replace(strComputerName, Chr(0), ""))
End Function

Public Sub WriteTxtLog(ByVal strText As String)
    '写入升级日志
    
    mobjText.WriteLine strText
End Sub

Public Sub OpenLogFile(ByVal blnPreUpgrade As Boolean)
    Dim strTxtFile  As String
    If blnPreUpgrade Then
        strTxtFile = gstrAppPath & "\ZLPerUpGradeList.Lst" '预升级日志
    Else
        strTxtFile = gstrAppPath & "\ZLUpGradeList.Lst" '升级日志
    End If
    
    '打开日志文件
    If FindFile(strTxtFile) = False Then
        mobjFile.CreateTextFile (strTxtFile)
    End If
    '--mobjFile.CreateTextFile (strTxtFile)
    
    Set mobjText = mobjFile.OpenTextFile(strTxtFile, ForWriting, True)
End Sub
Public Sub CloseLogFile()
    '关闭日志文件
    mobjText.Close
End Sub

Public Function GetCommpentVersion(ByVal strFile As String) As String

    '功能:获取指定控件的版本号
    '入参:
    '出参:
    '返回:成功,返回版本号,否则返回空
    '编制:刘兴洪
    '日期:2009-01-16 16:59:34

    Dim objFile As New FileSystemObject
    Dim strVer As String, varVersion As Variant
    
    On Error Resume Next
    '获取文件版本号
    strVer = objFile.GetFileVersion(strFile)
    If Err <> 0 Then
        Err.Clear
        GetCommpentVersion = ""
        Exit Function
    End If
    If Trim(strVer) <> "" Then
        varVersion = Split(strVer, ".")
        If UBound(varVersion) > 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(3)
        ElseIf UBound(varVersion) = 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(2)
        End If
    End If
    GetCommpentVersion = strVer
End Function

Private Function zlGetFileProcess(ByVal strFile As String, ByRef cllOutProcess As Collection) As Boolean

    '功能:获取指定文件的相关进程
    '入参:strFile-指定的DLL文件
    '出参:cllOutProcess-返回被引用的进程值
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-01-20 13:59:35

    Dim uProcess As PROCESSENTRY32, uMdlInfor As MODULEENTRY32
    Dim lngMdlProcess As Long, strExeName As String, lngSnapShot As Long, strDLLName As String
    
    On Error GoTo ErrHand:
    '创建进程快照
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot > 0 Then
      uProcess.lSize = Len(uProcess)
      If Process32First(lngSnapShot, uProcess) Then
        Do
          '获得进程的标识符
          strExeName = UCase(Left(Trim(uProcess.sExeFile), InStr(1, Trim(uProcess.sExeFile), vbNullChar) - 1))
          If strExeName Like "*" & UCase(strFile) & "*" Then
             '一般来说只有Exe文件才会存在
            On Error Resume Next
            cllOutProcess.Add Array(uProcess.lProcessId, strExeName, uProcess.lProcessId), "B" & uProcess.lProcessId
            If Err <> 0 Then
                cllOutProcess.Remove "B" & uMdlInfor.th32ProcessID
                cllOutProcess.Add Array(uProcess.lProcessId, strExeName, uProcess.lProcessId), "B" & uProcess.lProcessId
            End If
            On Error GoTo ErrHand:
          Else
                lngMdlProcess = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, uProcess.lProcessId)
                If lngMdlProcess > 0 Then
                    uMdlInfor.dwSize = Len(uMdlInfor)
                    If Module32First(lngMdlProcess, uMdlInfor) Then
                          Do
                                strDLLName = UCase(Left(Trim(uMdlInfor.szExePath), InStr(1, Trim(uMdlInfor.szExePath), vbNullChar) - 1))
                                If uProcess.lProcessId = uMdlInfor.th32ProcessID Then
                                    If strDLLName Like "*" & UCase(strFile) & "*" Then
                                        On Error Resume Next
                                        cllOutProcess.Add Array(uProcess.lProcessId, strExeName, uMdlInfor.th32ProcessID), "K" & uMdlInfor.th32ProcessID
                                        If Err <> 0 Then
                                            cllOutProcess.Remove "K" & uMdlInfor.th32ProcessID
                                            cllOutProcess.Add Array(uProcess.lProcessId, strExeName, uMdlInfor.th32ProcessID), "K" & uMdlInfor.th32ProcessID
                                        End If
                                        On Error GoTo ErrHand:
                                    End If
                                End If
                          Loop Until (Module32Next(lngMdlProcess, uMdlInfor) < 1)
                    End If
                    CloseHandle (lngMdlProcess)
                End If
            End If
        Loop Until (Process32Next(lngSnapShot, uProcess) < 1)
      End If
      CloseHandle (lngSnapShot)
    End If
    zlGetFileProcess = True
    Exit Function
ErrHand:
End Function
Private Function TerminatePID(ByVal lngPid As Long) As Boolean

    '功能:结束指定的进程
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-30 11:06:16

    Dim lngProcess As Long, pHandle As Long, ret As Long
    TerminatePID = False
    
    On Error GoTo ErrHand:
    pHandle = OpenProcess(SYNCHRONIZE, False, lngPid)
    lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPid)
    Call TerminateProcess(lngProcess, 1&)
    ret = WaitForSingleObject(pHandle, INFINITE)
    ret = CloseHandle(pHandle)
    TerminatePID = True
ErrHand:
End Function
 
Public Sub InitUpType()
'功能:初始升级方式信息

    On Error GoTo errH
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTemp As String
    
    strSQL = " Select 项目,内容 From zlregInfo where 项目= '升级类型'"
    Set rsTmp = OpenSQLRecord(strSQL, "升级类型")

    If rsTmp.EOF = False Then
        strTemp = NVL(rsTmp!内容, "0")
        If strTemp = "1" Then
             gintUpType = 1
        Else
             gintUpType = 0
        End If
    Else
        gintUpType = 0
    End If
    Exit Sub
errH:
    If Err Then
       gintUpType = 0
    End If
End Sub

Public Sub iniGatherTYpe()
'功能:初始收集方式信息

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTemp As String
    
    On Error GoTo errH
    strSQL = " Select 项目,内容 From zlregInfo where 项目= '收集方式'"
    Set rsTmp = OpenSQLRecord(strSQL, "收集方式")

    If rsTmp.EOF = False Then
        strTemp = NVL(rsTmp!内容, "0")
        If strTemp = "1" Then
             gintGatherTYpe = 1
        Else
             gintGatherTYpe = 0
        End If
    Else
        gintGatherTYpe = 0
    End If
    Exit Sub
errH:
    If Err Then
       gintGatherTYpe = 0
    End If
End Sub

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function GetSetupPath(ByVal strFileName As String, ByVal strPathSign As String, ByVal strFileType As String, ByVal strPath As String, ByVal str业务部件 As String) As String
    '功能:获取收集文件的完整路径
    '返回:返回完整的路径
    '如参 strFileName 文件名称
    '     strPathSign 文件安装路径
    '     strFileType 文件类型 0公共文件,1应用文件,2帮助文件,3其它文件,4三方控件
    '     stPath      当前所在的目录
    '     str业务部件 该部件的上层业务部件
    '编制:祝庆
    '日期:2010/12/10

    Dim strTemp As String '临时路径组合
    Dim strSystemDirectory As String '系统system32目录
    Dim strWinDirectory As String  'windows目录
    Dim blnExits        As Boolean
    Dim strTempProfession() As String
    Dim strTempFile     As String
    Dim i As Long
    
    On Error GoTo errH
    strSystemDirectory = GetWinSystemPath
    strWinDirectory = GetWinPath
    
    If strFileName = "" Then
        GetSetupPath = ""
        Exit Function
    End If
    
    If Len(strPathSign) = 0 Then
        Select Case strFileType
        Case "0" '公共
            strTemp = strPath & "\PUBLIC\" & strFileName
        Case "1" '应用
            strTemp = strPath & "\Apply\" & strFileName
        Case "2" '帮助
            strTemp = strWinDirectory & "\Help\" & strFileName
        Case "3" '其它
            strTemp = strPath & "\" & strFileName
        Case "4" '三方
            strTemp = ""
        Case "5"
            strPathSign = UCase(strPathSign)
            If (InStrRev(strPathSign, "[SYSTEM]", -1) > 0) Or (strPathSign = "") Then
                strTemp = strSystemDirectory & "\" & strFileName
            End If
            
            '新路径
            If InStrRev(strPathSign, "[PUBLIC]", -1) > 0 Then
                strTemp = strPath & "\PUBLIC\" & strFileName
            End If
            
        End Select
    Else
        strPathSign = UCase(strPathSign)
        If InStrRev(strPathSign, "[APPSOFT]", -1) > 0 Then
            strTemp = Replace(strPathSign, "[APPSOFT]", strPath)
            If Right(strTemp, 1) <> "\" Then
                strTemp = strTemp & "\" & strFileName
            Else
                strTemp = strTemp & strFileName
            End If
            
            '特殊处理批处理文件 "ZLAUTORUN.INI"
            If UCase(strFileName) = UCase("ZLAUTORUN.INI") Then
                strTemp = Replace(strTemp, ".INI", ".BAT")
            End If
        ElseIf InStrRev(strPathSign, "[SYSTEM]", -1) > 0 Then
            strTemp = Replace(strPathSign, "[SYSTEM]", strSystemDirectory)
            If Right(strTemp, 1) <> "\" Then
                strTemp = strTemp & "\" & strFileName
            Else
                strTemp = strTemp & strFileName
            End If
        ElseIf InStrRev(strPathSign, "[PUBLIC]", -1) > 0 Then
            strTemp = Replace(strPathSign, "[PUBLIC]", strPath & "\PUBLIC")
            If Right(strTemp, 1) <> "\" Then
                strTemp = strTemp & "\" & strFileName
            Else
                strTemp = strTemp & strFileName
            End If
        ElseIf InStrRev(strPathSign, "[HELP]", -1) Then
            strTemp = Replace(strPathSign, "[HELP]", strWinDirectory & "\Help")
            If Right(strTemp, 1) <> "\" Then
                strTemp = strTemp & "\" & strFileName
            Else
                strTemp = strTemp & strFileName
            End If
        Else '完整路径
            If Left(strFileName, 2) = "\\" Then
                strTemp = ""
            Else
                strTemp = Left(strPath, 1) & Right(strFileName, Len(strFileName) - 1)
            End If
        End If
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''判断业务部件标准
'''1.先检查业务部件在本机是否已经存在
'''1.1存在
'''    下载
'''1.2不存在
'''    检查它的业务部件是否载业务列表里
'''    有
'''        检查业务部件是否存在，只要有一个存在，就下载,否则不下载。
'''    无
'''        不下载,说明本机不需要使用这个部件。
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '判断应用部件是否安装
    Select Case strFileType
    Case 1 '应用部件
        If mobjFile.FileExists(strTemp) Then
            '下载
        Else
'      测试串：     str业务部件 = "ZL9BILLEDIT.OCX,ZL9COMLIB.DLL,ZL9DESIGN.OCX,ZL9FUNCTION.DLL,ZL9PRINTMODE.DLL,ZL9REPORT.DLL"
            If Len(str业务部件) > 0 Then
                If UCase(GetNoSuffixFilename(strFileName)) = UCase(str业务部件) Then GoTo goOn
                strTempProfession = Split(str业务部件, ",")
                For i = 0 To UBound(strTempProfession)
                    If UCase(strTempProfession(i)) = UCase("zlSvrStudio") Then '管理工具特殊处理
                        strTempFile = strPath & "\" & strTempProfession(i) & ".exe"
                    Else
                        strTempFile = strPath & "\Apply\" & strTempProfession(i) & ".dll"
                    End If
                    
                    If mobjFile.FileExists(strTempFile) Then
                        blnExits = True
                        Exit For
                    Else
                        If UCase(strFileName) = UCase(str业务部件) & ".DLL" Then
                            blnExits = True
                            Exit For
                        End If
                    End If
                Next
                
                If blnExits Then
                    '下载
                Else
                    strTemp = ""
                End If
            Else
                strTemp = ""
            End If
        End If
    End Select
goOn:
    
    GetSetupPath = strTemp
    Exit Function
errH:
    If Err Then
        GetSetupPath = ""
    End If
End Function

Private Function IsSourceCode() As Boolean
    '功能:确定是否源代码
    '返回:是原代码-true,不是源代码-false
    On Error Resume Next
    Debug.Print 1 / 0
    IsSourceCode = Err <> 0
End Function

Public Function FileTempDown(ByVal strTargetFile As String, strFileName As String, Optional ByRef strErrMsg As String = "下载完成！") As Boolean
    Dim strTempDirectored As String '临时文件目录
    Dim strPerTempDirectored As String '预升级临时文件目录
    Dim strPerFileName As String '预升级文件
    
    On Error GoTo errH
    If gblnPreUpgrade Then
        strTempDirectored = gstrPerTempPath
    Else
        strTempDirectored = gstrTempPath
    End If
           
    If mobjFile.FolderExists(strTempDirectored) = False Then
       Call mobjFile.CreateFolder(strTempDirectored)
    End If
    
    '定时正式升级流程
    If gblnOfficialUpgrade Then
        strPerFileName = gstrPerTempPath & strFileName
        If gbln预升完成 Then
            If mobjFile.FileExists(strPerFileName) Then
                strFileName = strTempDirectored & strFileName
    
                If mobjFile.FileExists(strFileName) Then
                   If FileSystem.GetAttr(strFileName) <> vbNormal Then
                        Call FileSystem.SetAttr(strFileName, vbNormal)
                   End If
                   Call mobjFile.DeleteFile(strFileName)
                End If
                
                Err.Clear
                mobjFile.CopyFile strPerFileName, strFileName, True
                If Err = 0 Then
                    '完成就退出,不在下载
                    FileTempDown = True
                    strErrMsg = "下载完成！"
                    Exit Function
                End If
            End If
        End If
    End If
    
    '普通升级流程
    If InStrRev(strFileName, ":") = 0 Then
        strFileName = strTempDirectored & strFileName
        
        If mobjFile.FileExists(strFileName) Then
           If FileSystem.GetAttr(strFileName) <> vbNormal Then
                Call FileSystem.SetAttr(strFileName, vbNormal)
           End If
           Call mobjFile.DeleteFile(strFileName)
        End If
    End If
   
    If gintUpType = 0 Then
        If mobjFile.FileExists(strTargetFile) = False Then
'''            '文件存在,则改属性
'''            If FileSystem.GetAttr(strTargetFile) <> vbNormal Then
'''                    FileSystem.SetAttr strTargetFile, vbNormal
'''            End If
'''        Else
            strErrMsg = "文件在服务器目录不存在!"
            FileTempDown = False
            Exit Function
        End If
        
        Err.Clear
        mobjFile.CopyFile strTargetFile, strFileName, True
        If Err <> 0 Then
'            MsgBox "无法下载升级列表文件" & strTargetFile & vbCrLf & "请确认共享服务器上是否存在!", vbInformation + vbDefaultButton1, "客户端自动升级"
            strErrMsg = "文件在服务器目录不存在!"
            FileTempDown = False
            Exit Function
        Else
            FileTempDown = True
            strErrMsg = "下载完成！"
        End If
    Else
        If FtpDownFile(strTargetFile, strFileName) = False Then
'            MsgBox "无法下载升级列表文件" & strTargetFile & vbCrLf & "请确认FTP服务器上是否存在!", vbInformation + vbDefaultButton1, "客户端自动升级"
            strErrMsg = "文件在服务器目录不存在!"
            FileTempDown = False
            Exit Function
        Else
             FileTempDown = True
             strErrMsg = "下载完成！"
        End If
    End If
    Exit Function
errH:
    If Err Then
        FileTempDown = False
    End If
End Function

Public Function FileDeCompression(strTempFile As String, Optional ByRef strErrMsg As String = "解压缩完成！") As Boolean
    '功能:解压缩文件到本地,并删除压缩是文件!
    Dim strDeCompTxt As String
    Dim strSource    As String
    Dim strTemp  As String
    Dim RetVal       As Long
    Dim i            As Integer
    
    On Error GoTo errH
    i = InStrRev(strTempFile, "\", -1)
    If i > 0 Then
        strSource = Left(strTempFile, i)
    End If
    
    strTemp = Left(strTempFile, Len(strTempFile) - 3)
    
    strDeCompTxt = DeCompressionCmd(strSource, strTempFile)
    If strDeCompTxt <> "" Then
        Call GetCmdTxt(strDeCompTxt)
'        RetVal = Shell(strDeCompTxt, vbHide)
       
        On Error Resume Next
        Call Kill(strTempFile)
        strTempFile = strTemp
        strErrMsg = "解压缩完成！"
        FileDeCompression = True
    Else
        strErrMsg = "解压缩失败!"
        FileDeCompression = False
    End If
    Exit Function
errH:
    If Err Then
        FileDeCompression = False
    End If
End Function

Public Function GetErrParameter(ByVal intParameterNum As Integer) As String
'功能:根据条件获取参数值
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "select 参数值 from ZlOptions where 参数号=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "参数号", intParameterNum)

    If rsTmp.RecordCount = 1 Then
         GetErrParameter = NVL(rsTmp!参数值, "0")
    Else
         GetErrParameter = "0"
    End If
    Exit Function
errH:
    If Err Then
        GetErrParameter = "0"
    End If
End Function

Public Sub SaveErrLog(ByVal strMsgInfo As String)
'功能：将错误信息写入数据库错误日志
    Dim strSQL As String
        On Local Error Resume Next
        Dim bytErrType As Byte, lngErrNum As Long
        bytErrType = 4 '客户端升级错误
        lngErrNum = 0  '错误号
        
        
        strSQL = "Insert Into zlErrorLog(会话号,用户名,工作站,时间,类型,错误序号,错误信息)" & _
            " Select SID,USER,MACHINE,SYSDATE," & bytErrType & "," & lngErrNum & "," & _
            AdjustStr(strMsgInfo) & " From v$Session Where AUDSID=UserENV('SessionID')"
        gcnOracle.Execute strSQL
End Sub

Private Function AdjustStr(Str As String) As String
'功能：将含有"'"符号的字符串调整为Oracle所能识别的字符常量
'说明：自动(必须)在两边加"'"界定符。

    Dim i As Long, strTmp As String
    
    If InStr(1, Str, "'") = 0 Then AdjustStr = "'" & Str & "'": Exit Function
    
    For i = 1 To Len(Str)
        If Mid(Str, i, 1) = "'" Then
            If i = 1 Then
                strTmp = "CHR(39)||'"
            ElseIf i = Len(Str) Then
                strTmp = strTmp & "'||CHR(39)"
            Else
                strTmp = strTmp & "'||CHR(39)||'"
            End If
        Else
            If i = 1 Then
                strTmp = "'" & Mid(Str, i, 1)
            ElseIf i = Len(Str) Then
                strTmp = strTmp & Mid(Str, i, 1) & "'"
            Else
                strTmp = strTmp & Mid(Str, i, 1)
            End If
        End If
    Next
    AdjustStr = strTmp
End Function

Private Sub CreatePath(ByVal strPath As String)
    Dim i As Integer, j As Integer
    Dim strTemp As String
    
    On Error Resume Next
    i = InStrRev(strPath, "\", -1)
    If i > 0 Then
        strTemp = Left(strPath, i)
        If mobjFile.FolderExists(strTemp) = False Then
           Call MakeDirX(strTemp)
        End If
    End If
End Sub

Public Function MakeDirX(ByVal strPath As String) As Boolean
'检查文件路径是否存在，如不存在就创建
  Dim strPan As String, mPath As Long
  Dim ret As Integer, ret0 As Integer
  
  ret = InStr(strPath, "\\")
  If ret > 0 Then
    ret0 = ret + 2
  Else: ret0 = 1
  End If
  On Error Resume Next
  Dim WFD As WIN32_FIND_DATA
  Do
    Err.Clear
    ret = InStr(ret0, strPath, "\", vbTextCompare)
    If ret > 0 Then
      strPan = Left(strPath, ret - 1)
      mPath = FindFirstFile(strPan & "\*.*", WFD)
      If mPath = INVALID_HANDLE_VALUE Then
        MkDir strPan
      End If
      ret0 = ret + 1
    End If
  Loop Until ret = 0
  If Err = 76 Then
    Exit Function
  End If
  MakeDirX = True
End Function

Public Function GetTmpPath() As String
    Dim tmpBuffer As String
    tmpBuffer = String(255, Chr(0))
    GetTempPath 256, tmpBuffer
    GetTmpPath = Trim(Left(tmpBuffer, InStr(1, tmpBuffer, Chr(0)) - 1))
End Function

Public Function CheckSysFile(ByVal strTargetFile As String) As Boolean
    Dim strTempFileName As String
    Dim i As Integer
    If InStrRev(strTargetFile, GetWinSystemPath, -1) > 0 Then
        i = InStrRev(strTargetFile, "\")
        If i > 0 Then
            strTempFileName = Right(strTargetFile, Len(strTargetFile) - i)
            If UCase(Left(strTempFileName, 2)) <> "ZL" Then
                CheckSysFile = True
                Exit Function
            End If
        End If
    End If
End Function

Private Function GetNoSuffixFilename(ByVal strFileName As String) As String
    '去掉后缀的文件名
    Dim i As Integer
    Dim strTmp As String
    If strFileName <> "" Then
        i = InStrRev(strFileName, ".", -1)
        If i > 0 Then
            strTmp = Left(strFileName, i - 1)
        End If
    End If
    GetNoSuffixFilename = strTmp
End Function

Public Function GetAdmin() As Boolean
    '判断是否具有过来员权限,生成测试文件zlTestAdmin拷贝自身到system32下
    Dim strAppPath As String
    
    On Error Resume Next
    strAppPath = App.Path & "\zlTestAdmin.txt"
    
    
    Open strAppPath For Output As #1
    Print #1, Now & "   检测管理员权限"
    Close #1
    FileCopy strAppPath, GetWinSystemPath & "\zlTestAdmin.txt"

    If Err.Number = 75 Then
        GetAdmin = False
        'MsgBox "没有管理员权限"
    ElseIf Dir(GetWinSystemPath & "\zlTestAdmin.txt", vbNormal) <> "" Then
        GetAdmin = True
        Call Kill(GetWinSystemPath & "\zlTestAdmin.txt")
        'MsgBox "有管理员权限"
    Else
        GetAdmin = False
    End If
    
    '删除测试文件zlTestAdmin
    Call Kill(strAppPath)
End Function

'系统管理员密码解密方法
Public Function decipher(stext As String)      '密码解密程序
    Const min_asc = 32 '最小ASCII码
    Const max_asc = 126 '最大ASCII码 字符
    Const num_asc = max_asc - min_asc + 1
    Dim offset As Long
    Dim strlen As Integer
    Dim i As Integer
    Dim ch As Integer
    Dim ptext As String
    offset = 123
    Rnd (-1)
    Randomize (offset)
    strlen = Len(stext)
    For i = 1 To strlen
       ch = Asc(Mid(stext, i, 1)) '取字母转变成ASCII码
       If ch >= min_asc And ch <= max_asc Then
           ch = ch - min_asc
           offset = Int((num_asc + 1) * Rnd())
           ch = ((ch - offset) Mod num_asc)
           If ch < 0 Then
               ch = ch + num_asc
           End If
           ch = ch + min_asc
           ptext = ptext & Chr(ch)
       End If
    Next i
    decipher = ptext
End Function

Public Sub SaveClientLog(ByVal strMsgInfo As String)
'功能：将错误信息写入数据库错误日志
    Dim strSQL As String
    On Local Error Resume Next

    strSQL = "Insert Into zltools.zlClientUpdatelog(工作站,处理日期,内容)" & _
        " Select TERMINAL,SYSDATE," & _
        AdjustStr(strMsgInfo) & " From v$Session Where AUDSID=UserENV('SessionID')"
    gcnOracle.Execute strSQL
       
End Sub

Public Sub UpdateCondition(ByVal intMode As Integer)
'功能：升级情况
    Dim strSQL As String
    On Local Error Resume Next

    If intMode = 1 Then
        strSQL = "Zl_Zlclients_Control(15,'" & gstrComputerName & "')"
    Else
        strSQL = "Zl_Zlclients_Control(16,'" & gstrComputerName & "')"
    End If
    Call ExecuteProcedure(strSQL, "UpdateCondition")
End Sub



