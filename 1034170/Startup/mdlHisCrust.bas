Attribute VB_Name = "mdlHisCrust"
Option Explicit

'分析本机配置相关API
'----------------------------------------------------------------------------------------------------
'Window版本函数
'win2000 以下版本
Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

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
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
'wSuiteMask
Private Const VER_SUITE_BACKOFFICE = &H4                'Microsoft BackOffice components are installed.
Private Const VER_SUITE_BLADE = &H400                   'Windows Server 2003, Web Edition is installed.
Private Const VER_SUITE_COMPUTE_SERVER = &H4000         'Windows Server 2003, Compute Cluster Edition is installed.
Private Const VER_SUITE_DATACENTER = &H80               'Windows Server 2008 Datacenter, Windows Server 2003, Datacenter Edition, or Windows 2000 Datacenter Server is installed.
Private Const VER_SUITE_ENTERPRISE = &H2                'Windows Server 2008 Enterprise, Windows Server 2003, Enterprise Edition, or Windows 2000 Advanced Server is installed. Refer to the Remarks section for more information about this bit flag.
Private Const VER_SUITE_EMBEDDEDNT = &H40               'Windows XP Embedded is installed.
Private Const VER_SUITE_PERSONAL = &H200                'Windows Vista Home Premium, Windows Vista Home Basic, or Windows XP Home Edition is installed.
Private Const VER_SUITE_SINGLEUSERTS = &H100            'Remote Desktop is supported, but only one interactive session is supported. This value is set unless the system is running in application server mode.
Private Const VER_SUITE_SMALLBUSINESS = &H1             'Microsoft Small Business Server was once installed on the system, but may have been upgraded to another version of Windows. Refer to the Remarks section for more information about this bit flag.
Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED = &H20 'Microsoft Small Business Server is installed with the restrictive client license in force. Refer to the Remarks section for more information about this bit flag.
Private Const VER_SUITE_STORAGE_SERVER = &H2000         'Windows Storage Server 2003 R2 or Windows Storage Server 2003is installed.
Private Const VER_SUITE_TERMINAL = &H10                 'Terminal Services is installed. This value is always set.
                                                        'If VER_SUITE_TERMINAL is set but VER_SUITE_SINGLEUSERTS is not set, the system is running in application server mode.
Private Const VER_SUITE_WH_SERVER = &H8000              'Windows Home Server is installed.
'wProductType
Private Const VER_NT_DOMAIN_CONTROLLER = &H2            'The system is a domain controller and the operating system is Windows Server 2012 , Windows Server 2008 R2, Windows Server 2008, Windows Server 2003, or Windows 2000 Server.
Private Const VER_NT_SERVER = &H3                       'The operating system is Windows Server 2012, Windows Server 2008 R2, Windows Server 2008, Windows Server 2003, or Windows 2000 Server.
                                                        'Note that a server that is also a domain controller is reported as VER_NT_DOMAIN_CONTROLLER, not VER_NT_SERVER.
Private Const VER_NT_WORKSTATION = &H1                  'The operating system is Windows 8, Windows 7, Windows Vista, Windows XP Professional, Windows XP Home Edition,
'dwPlatformId
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
'GetSystemMetrics
Private Const SM_TABLETPC = 86                          'Windows XP Tablet PC Edition
Private Const SM_MEDIACENTER = 87                       'Windows XP Media Center Edition
Private Const SM_STARTER = 88                           'Windows XP Starter Edition
Private Const SM_SERVERR2 = 89                          'Windows Server 2003 R2
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Type SYSTEM_INFO
'    dwOemID As Long
    wProcessorArchitecture As Integer
    wReserved As Integer
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
'wProcessorArchitecture
Private Const PROCESSOR_ARCHITECTURE_AMD64 = 9          'x64 (AMD Or Intel)
Private Const PROCESSOR_ARCHITECTURE_ARM = 5            'ARM
Private Const PROCESSOR_ARCHITECTURE_IA64 = 6           'Intel Itanium - based
Private Const PROCESSOR_ARCHITECTURE_INTEL = 0          'x86
Private Const PROCESSOR_ARCHITECTURE_UNKNOWN = &HFFFF   'Unknown architecture.
Private Const PROCESSOR_INTEL_386 = 386
Private Const PROCESSOR_INTEL_486 = 486
Private Const PROCESSOR_INTEL_PENTIUM = 586
Private Const PROCESSOR_INTEL_IA64 = 2200
Private Const PROCESSOR_AMD_X8664 = 8664
Private Const PROCESSOR_MIPS_R4000 = 4000      ' incl R4101 & R3910 for Windows CE
Private Const PROCESSOR_ALPHA_21064 = 21064
Private Const PROCESSOR_PPC_601 = 601
Private Const PROCESSOR_PPC_603 = 603
Private Const PROCESSOR_PPC_604 = 604
Private Const PROCESSOR_PPC_620 = 620
Private Const PROCESSOR_HITACHI_SH3 = 10003    ' Windows CE
'获取内存
Private Type MEMORYSTATUS  'win2000及以下版本
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long

End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Type MEMORYSTATUSEX
        dwLength       As Long
        dwMemoryLoad   As Long
        ullTotalPhys   As Currency
        ullAvailPhys   As Currency
        ullTotalPageFile   As Currency
        ullAvailPageFile   As Currency
        ullTotalVirtual    As Currency
        ullAvailVirtual    As Currency
        ullAvailExtendedVirtual   As Currency
End Type
Private Declare Function GlobalMemoryStatusEx Lib "kernel32.dll" (ByRef lpBuffer As MEMORYSTATUSEX) As Long
'取硬盘大小
Private Const DRIVE_FIXED = 3
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const STRSPLIT As String = "♂♂"

'API错误信息获取
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
Private Const ERROR_EXTENDED_ERROR          As Long = 1208
'文件描述信息判断
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (ByVal pBlock As Long, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
'Public Const FVN_Comments           As String = "Comments"          '注释
'Public Const FVN_InternalName       As String = "InternalName"      '内部名称
'Public Const FVN_ProductName        As String = "ProductName"       '产品名
'Public Const FVN_CompanyName        As String = "CompanyName"       '公司名
'Public Const FVN_ProductVersion     As String = "ProductVersion"    '产品版本
'Public Const FVN_FileDescription    As String = "FileDescription"   '文件描述
'Public Const FVN_OriginalFilename   As String = "OriginalFilename"  '原始文件名
'Public Const FVN_FileVersion        As String = "FileVersion"       '文件版本
'Public Const FVN_SpecialBuild       As String = "SpecialBuild"      '特殊编译号
'Public Const FVN_PrivateBuild       As String = "PrivateBuild"      '私有编译号
'Public Const FVN_LegalCopyright     As String = "LegalCopyright"    '合法版权
'Public Const FVN_LegalTrademarks    As String = "LegalTrademarks"   '合法商标
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'hModule：一个模块的句柄。可以是一个DLL模块，或者是一个应用程序的实例句柄。如果该参数为NULL，该函数返回该应用程序全路径?
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
'电脑名称(ComputerName)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'临时IP获取
Private Const MAX_IP = 5                                                    'To make a buffer... i dont think you have more than 5 ip on your pc..
Private Type IPINFO
    dwAddr As Long                                                          ' IP address
    dwIndex As Long                                                         ' interface index
    dwMask As Long                                                          ' subnet mask
    dwBCastAddr As Long                                                     ' broadcast address
    dwReasmSize  As Long                                                    ' assembly size
    unused1 As Integer                                                      ' not currently used
    unused2 As Integer                                                      '; not currently used
End Type
Private Type MIB_IPADDRTABLE
    dEntrys As Long                                                         'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO                                               'array of IP address entries
End Type
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public gstrExeFile      As String '调用登录部件的EXE路径
Public gstrSetupPath    As String 'APPSOFT路径
Public glnghInstance    As Long
Public gblnTimer            As Boolean  '是否定时器触发的客户端更新检查

Public Function CheckAllowByTerminal() As Boolean
'功能:检查是否允许使用本工作站,以及进行当前工作站信息的登记
'     判断是否允许该工作站使用程序；
'     如果需要替换本地参数，则执行替换操作；如果需要升级，则调用外壳程序，并关闭退出
'返回:成功,返回true,否则返回False
'警告：由于还没有初始化公共部件的连接对象，该函数中不能使用公共部件中的数据库访问方法

    Dim rsTmp As ADODB.Recordset, strSQL As String, strRowID As String '客户端的ROWID
    Dim strComuterInfo As String, arrComputer As Variant, strComputerName As String, strIpAddress As String
    Dim strTmp As String, arrTmp As Variant, i As Integer
    Dim bln检查站点 As Boolean, lng有站点 As Long, bln空站点 As Boolean, bln多站点 As Boolean
    Dim str站点       As String, str站点编号 As String, str名称 As String, str缺省部门
    Dim blnAllow As Boolean, blnUpdate As Boolean
    Dim int服务器编号 As Integer, int启用视频源 As Integer, int连接数 As Integer, int升级标志 As Integer
    
'    Call SQLTest(App.EXEName, "mdlHisCrust", "新版电子病历自动升级检查")
    Call UpdateEmrInterface '新版电子病历自动升级
'    Call SQLTest

    strIpAddress = IP '以oracle连接的IP地址为主
    strComputerName = ComputerName
    '检查是否有重名机器
    If CheckRepeatLogin(strIpAddress) = True Then
        CheckAllowByTerminal = False
        Exit Function
    End If
    '判断是否允许使用
    strComuterInfo = AnalyseConfigure
    arrComputer = Split(strComuterInfo, STRSPLIT)
    '1.以站点名检查
    If Err.Number <> 0 Then Err.Clear
    On Error Resume Next
    strSQL = "Select Rowid as ID,站点,部门,Nvl(禁止使用,0) as 允许,Nvl(升级标志,0) as 升级,Nvl(收集标志,0) as 收集,连接数,启用视频源 From zlClients Where 工作站=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "检查工作站-以站点为主", strComputerName)
    '可能由于未授权等原因，导致查询出错，此时弹出提示禁止登录
    If rsTmp Is Nothing Then
        MsgBox Err.Description & vbNewLine & "不能正常访问系统，请您联系系统管理员重新进行角色授权！", vbInformation, gstrSysName
        Exit Function
    End If
    '2.未发现此站点,则以IP方式查找，但只有一个时才更新计算名
    If rsTmp.EOF Then
        strSQL = "Select Rowid as ID,站点,部门, Nvl(禁止使用,0) as 允许,Nvl(升级标志,0) as 升级,Nvl(收集标志,0) as 收集,连接数,启用视频源 From zlClients Where IP=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "检查工作站-以站点为主", strIpAddress)
        If rsTmp.RecordCount > 1 Then
            '大于两个以上,则加CPU,内存,硬盘为限制条件.
            strSQL = "" & _
                "   Select Rowid as ID,站点,部门,Nvl(禁止使用,0) as 允许,Nvl(升级标志,0) as 升级,Nvl(收集标志,0) as 收集,连接数,启用视频源 " & _
                "   From zlClients Where IP=[1] and CPU=[2] and  内存=[3] and 硬盘=[4]"
            Set rsTmp = OpenSQLRecord(strSQL, "检查工作站-以站点为主", strIpAddress, CStr(arrComputer(2)), CStr(arrComputer(3)), CStr(arrComputer(4)))
        End If
    End If
    bln检查站点 = True
    '如果还存在多个,则可能存在IP冲突的情况,因此不能判定需要更新相关的站点.只能当成新的站点上传
    If rsTmp.RecordCount > 1 Or rsTmp.EOF Then
        strRowID = ""
    Else '表示更新相关的信息
        strRowID = NVL(rsTmp!id)
        int启用视频源 = Val(NVL(rsTmp!启用视频源))
        '升级后登陆,不在让用户选择,直接读取
        If gstrCommand <> "" Then
            '新方法
            If InStr(gstrCommand, "ZLHISCRUSTCALL=1") > 0 And InStr(gstrCommand, "USER=") > 0 And InStr(gstrCommand, "PASS=") > 0 Then
                bln检查站点 = False
                str站点编号 = NVL(rsTmp!站点)
                gobjRelogin.DeptName = NVL(rsTmp!部门)
            '老的判断方法
            ElseIf InStrRev(gstrCommand, "/", -1) > 0 And InStrRev(gstrCommand, ",", -1) = 0 Then
                bln检查站点 = False
                str站点编号 = NVL(rsTmp!站点)
                gobjRelogin.DeptName = NVL(rsTmp!部门)
            End If
        End If
        blnAllow = Val(rsTmp!允许 & "") = 0
        int连接数 = Val(rsTmp!连接数 & "")  '0-表示无限制
        blnUpdate = Val(rsTmp!升级 & "") = 1
        If Not blnUpdate Then blnUpdate = Val(rsTmp!收集 & "") = 1
    End If

    If bln检查站点 Then
        strSQL = "Select b.名称, a.站点, a.缺省" & vbNewLine & _
                "From (Select c.站点, b.缺省" & vbNewLine & _
                "       From 上机人员表 a, 部门人员 b, 部门表 c" & vbNewLine & _
                "       Where a.人员id = b.人员id And b.部门id = c.Id And a.用户名 = Upper([1])) a, Zlnodelist b" & vbNewLine & _
                "Where a.站点 = b.编号(+)" & vbNewLine & _
                "Order By 站点"
        Set rsTmp = OpenSQLRecord(strSQL, "检查并确定所属院区", gobjRelogin.DBUser)
        If rsTmp Is Nothing Then
            MsgBox Err.Description & vbNewLine & "不能正常访问系统，请您联系系统管理员重新进行角色授权！", vbInformation, gstrSysName
            Exit Function
        End If
        Do While Not rsTmp.EOF
            If NVL(rsTmp!站点, "") <> "" Then
                str站点 = str站点 & "," & NVL(rsTmp!站点, "")
                str名称 = str名称 & "," & NVL(rsTmp!名称)
                lng有站点 = lng有站点 + 1
            Else
                bln空站点 = True
            End If
            If NVL(rsTmp!缺省, "0") = 1 Then
                str缺省部门 = NVL(rsTmp!名称)
            End If
            rsTmp.MoveNext
        Loop
        '如果当前登录人员所属部门都没有设置站点，则不作处理。在查找该院是否启动了站点控制!
        If str站点 = "" Or (bln空站点 And lng有站点 <> 1) Then
            '独立安装新版LIS时也需要按仪器读取站点
            strTmp = GetLISStation()
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ";")
                str站点 = arrTmp(0)
                str名称 = arrTmp(1)
            Else
                str站点 = "": str名称 = ""
                strSQL = "select distinct (A.站点),B.名称 from 部门表 A,zlNodeList B where A.站点=B.编号 And A.站点 is not null order by A.站点"
                Set rsTmp = OpenSQLRecord(strSQL, "检查是否启动站点控制")
                If Not rsTmp Is Nothing Then
                    Do While Not rsTmp.EOF
                        If NVL(rsTmp!站点, "") <> "" Then
                            str站点 = str站点 & "," & NVL(rsTmp!站点, "")
                            str名称 = str名称 & "," & NVL(rsTmp!名称)
                        End If
                        rsTmp.MoveNext
                    Loop
                End If
            End If
        End If
        If str站点 <> "" Then
            str站点 = Mid(str站点, 2)
            str名称 = Mid(str名称, 2)
            arrTmp = Split(str站点, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                If i = LBound(arrTmp) Then
                    str站点编号 = arrTmp(i)
                Else
                    If str站点编号 <> arrTmp(i) Then
                        bln多站点 = True
                        Exit For
                    End If
                End If
            Next
            If bln多站点 Then '提示用户选择当前计算机位置所在的部门。
                str站点编号 = GetSetting("ZLSOFT", "私有模块\" & gobjRelogin.DBUser & "\" & App.ProductName & "\" & App.EXEName, "当前站点选择", "")
                Call frmSelClient.ShowEdit(str站点, str名称, str站点编号)
                str站点编号 = IIf(frmSelClient.gstr站点 = "无", "", frmSelClient.gstr站点)
                gobjRelogin.DeptName = frmSelClient.gstrCur站点
                Call SaveSetting("ZLSOFT", "私有模块\" & gobjRelogin.DBUser & "\" & App.ProductName & "\" & App.EXEName, "当前站点选择", str站点编号)
            End If
        End If
    End If
    gobjRelogin.NodeNo = IIf(str站点编号 <> "", str站点编号, "-")
    If gobjRelogin.DeptName = "" Then gobjRelogin.DeptName = str缺省部门
    If strRowID = "" Then '新增的工作站，还没有该工作站的数据，上传（IP、机器名、CPU、内存、硬盘、操作系统）
        int服务器编号 = GetDefaultFileServer
        If int服务器编号 = -1 Then '获取默认服务器失败，则不升级，恢复服务器编号的初始值
            int服务器编号 = 0
            int升级标志 = 0
        Else
            int升级标志 = 1
        End If
        strSQL = "Zl_Zlclients_Set(0,Null,'" & strComputerName & "','" & strIpAddress & "','" & arrComputer(2) & "','" & arrComputer(3) & _
                    "','" & arrComputer(4) & "','" & arrComputer(5) & "','" & gobjRelogin.DeptName & "',Null,Null," & int服务器编号 & "," & int升级标志 & _
                    ",0,'" & str站点编号 & "',0,Null,Null," & int启用视频源 & ")"
        ExecuteProcedure strSQL, "新增工作站"
        '新增客户端不能升级则直接退出
        If int升级标志 = 0 Then
            CheckAllowByTerminal = True
            Exit Function
        End If
        blnUpdate = True
    Else
        strSQL = "Zl_Zlclients_Set(1,'" & strRowID & "','" & strComputerName & "','" & strIpAddress & "','" & arrComputer(2) & "','" & arrComputer(3) & _
                    "','" & arrComputer(4) & "','" & arrComputer(5) & "','" & gobjRelogin.DeptName & "',Null,Null,Null,Null," & int连接数 & ",'" & str站点编号 & "',0,Null,Null," & int启用视频源 & ")"
        '需要更新相关的站点信息
        ExecuteProcedure strSQL, "更新工作站"
        If Not blnAllow Then
            MsgBox "该工作站已被管理员禁用！", vbInformation, gstrSysName
            Exit Function
        End If
        '连接数检查限制
        If int连接数 > 0 Then
            strSQL = "Select SID From gv$Session Where Upper(PROGRAM) Like 'ZL%.EXE' And Status<>'KILLED' And MACHINE=(Select Max(MACHINE) From v$Session Where AUDSID=UserENV('SessionID'))"
            Set rsTmp = OpenSQLRecord(strSQL, "检查连接数量")
            If rsTmp.RecordCount > int连接数 Then
                MsgBox "当前工作站最多只允许 " & int连接数 & " 个登录连接，当前已经有 " & rsTmp.RecordCount - 1 & " 个连接。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    On Error GoTo Errhand
AutoUpGrude:      '执升升级程序
    If blnUpdate Then
        blnAllow = UpdateZLHIS(strComputerName)
    End If
    CheckAllowByTerminal = blnAllow
    Exit Function
Errhand:
    MsgBox "升级检查出现错误：" & Err.Description & "，请您联系系统管理员进行解决！", vbInformation, gstrSysName
End Function

Public Function StartHisCrust(ByVal str升级程序 As String, ByVal strJobName As String, Optional ByVal lngWait As Long, Optional ByVal strPass As String) As Boolean
'功能：调用自动升级外壳
'参数：str升级程序=可以直接传完成文件路径，也可以传文件名
'      strJobName=任务名称，或者调用程序名
'      lngWait=正式升级时，等待的N分钟后才正式升级
'返回：是否成功
    Dim strUP As String
    Dim strUPFile  As String, strFileName As String
    Dim strConnString As String, lngErr As Long
    Dim objFile As New FileSystemObject
    Dim strCheck As String, strCommand As String
    Dim strEXEName As String
    
    
    On Error Resume Next
    If gstrExeFile <> "" Then
        strEXEName = objFile.GetFileName(gstrExeFile)
    Else
        strEXEName = strJobName
    End If
    
    If objFile.GetDriveName(str升级程序) = "" Then
        strUPFile = gstrSetupPath & "\" & str升级程序
    Else
        strUPFile = str升级程序
        strFileName = objFile.GetFileName(str升级程序)
    End If
    If Not objFile.FileExists(strUPFile) Then
        MsgBox "没有找到客户端自动升级工具" & strFileName & "，请与系统管理员联系。", vbExclamation, gstrSysName
        Exit Function
    End If
    If IsDesinMode Then
        '组装命令行，以及生成命令行校验位
        strCommand = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=" & gobjRelogin.ServerName & _
                                   """;Persist Security Info=True;User ID=" & gobjRelogin.InputUser & ";Password=HIS;Data Provider=MSDASQL"
    Else
        '组装命令行，以及生成命令行校验位
        strCommand = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=" & gobjRelogin.ServerName & _
                                   """;Persist Security Info=True;User ID=" & gobjRelogin.InputUser & ";Password=" & strPass & ";Data Provider=MSDASQL"
    End If
    strCheck = "CMDCHECK:1" & "," & Len(strCommand)
    strCommand = strCommand & "||0"
    strCheck = strCheck & "," & Len(strCommand)
    strCommand = strCommand & "||" & strEXEName
    strCheck = strCheck & "," & Len(strCommand)
    strCommand = strCommand & "||" & CStr(gstrCommand)
    strCheck = strCheck & "," & Len(strCommand)
    strCommand = strCommand & "||" & "USER=" & gobjRelogin.InputUser & " PASS=" & gobjRelogin.InputPwd
    strCheck = strCheck & "," & Len(strCommand)
    If lngWait <> 0 Then
        strCommand = strCommand & "||W:" & lngWait
        strCheck = strCheck & "," & Len(strCommand)
    End If
    strCommand = strCommand & "||" & strCheck
    lngErr = Shell(strUPFile & " " & strCommand, vbNormalFocus)
    StartHisCrust = True
    If lngErr = 0 Then
        MsgBox "无法启动部件升级进程，请使用操作系统管理员身份启动程序。", vbInformation, gstrSysName
    End If
End Function

Private Function AnalyseConfigure() As String
    '编写人:朱玉宝 2003-03-09
    '功能:分析出本机的配置（IP、机器名、CPU、内存、硬盘、操作系统）
    Dim strCPU As String           'CPU
    Dim strMemory As String        '内存
    Dim strOS As String            '操作系统
    Dim strComputerName As String  '计算机名
    Dim strHD As String            '硬盘
    Dim strIp As String            'IP地址
    Dim verinfo As OSVERSIONINFOEX
    Dim sysinfo As SYSTEM_INFO
    Dim memsts As MEMORYSTATUS
    Dim memstsex As MEMORYSTATUSEX
    Dim lngmemory As Long
    Dim curMemory As Currency
    
    strIp = LoacalIP
    '获取计算机名
    strComputerName = ComputerName
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
    '先判断系统是否为win2000及以下
    '如果是Windows2000或以下版本，则用GlobalMemoryStatus取
    verinfo.dwOSVersionInfoSize = Len(verinfo) 'should be 148/156
    If GetVersionEx(verinfo) = 0 Then 'try win2000 version
        GlobalMemoryStatus memsts
        lngmemory = memsts.dwTotalPhys
        strMemory = Format$(lngmemory& \ 1024 \ 1024, "###,###,###") + "M"
    Else
        memstsex.dwLength = Len(memstsex)
        GlobalMemoryStatusEx memstsex
        curMemory = memstsex.ullTotalPhys
        strMemory = CStr(Int(curMemory * 10000 / 1024 ^ 2)) & "M"
    End If
    AnalyseConfigure = strIp & STRSPLIT & strComputerName & STRSPLIT & strCPU & _
                       STRSPLIT & strMemory & STRSPLIT & strHD & STRSPLIT & strOS
End Function

Private Function AnalyseHardDisk() As String
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
    Dim strOS As String
    Dim sysinfo As SYSTEM_INFO
    'OSVERSIONINFO
    'Operating system    Version number  dwMajorVersion  dwMinorVersion  Other
    'Windows 10                 10.0*       10                  0   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2016        10.0*       10                  0   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows 8.1                6.3*        6                   3   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2012 R2     6.3*        6                   3   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows 8                  6.2         6                   2   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2012        6.2         6                   2   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows 7                  6.1         6                   1   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2008 R2     6.1         6                   1   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows Server 2008        6.0         6                   0   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows Vista              6.0         6                   0   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2003 R2     5.2         5                   2   GetSystemMetrics(SM_SERVERR2) != 0
    'Windows Server 2003        5.2         5                   2   GetSystemMetrics(SM_SERVERR2) == 0
    'Windows XP                 5.1         5                   1   Not applicable
    'Windows 2000               5.0         5                   0   Not applicable
    'OSVERSIONINFOEX
    'Operating system    Version number  dwMajorVersion  dwMinorVersion  Other
    'Windows 10                 10.0*       10                  0   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2016        10.0*       10                  0   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows 8.1                6.3*        6                   3   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2012 R2     6.3*        6                   3   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows 8                  6.2         6                   2   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2012        6.2         6                   2   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows 7                  6.1         6                   1   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2008 R2     6.1         6                   1   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows Server 2008        6.0         6                   0   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows Vista              6.0         6                   0   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2003 R2     5.2         5                   2   GetSystemMetrics(SM_SERVERR2) != 0
    'Windows Home Server        5.2         5                   2   OSVERSIONINFOEX.wSuiteMask & VER_SUITE_WH_SERVER
    'Windows Server 2003        5.2         5                   2   GetSystemMetrics(SM_SERVERR2) == 0
    'Windows XP Professional x64 Edition 5.2    5               2   (OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION) && (SYSTEM_INFO.wProcessorArchitecture==PROCESSOR_ARCHITECTURE_AMD64)
    'Windows XP                 5.1         5                   1   Not applicable
    'Windows 2000               5.0         5                   0   Not applicable
    '如果是Windows2000或以下版本，则用新API再取一次
    myOS.dwOSVersionInfoSize = Len(myOS) 'should be 148/156
    If GetVersionEx(myOS) = 0 Then 'try win2000 version
        myOS.dwOSVersionInfoSize = 148 'if fails,ignore reserved data
        If GetVersionEx(myOS) = 0 Then
            GetVersionInfo = "Windows (Unknown)"
            Exit Function
        End If
    Else
        bExInfo = True
    End If
    ' 获得CPU类型
    GetSystemInfo sysinfo
    With myOS
        Select Case .dwMajorVersion
            Case 3
                strOS = "Windows NT 3.1"
            Case 4
                Select Case .dwMinorVersion
                    Case 0
                        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
                            strOS = "Windows NT 4.0" '1996年7月发布
                        Else
                            strOS = "Windows 95"
                        End If
                    Case 10
                        strOS = "Windows 98"
                    Case 90
                        strOS = "Windows Me"
                End Select
            Case 5
                Select Case .dwMinorVersion
                    Case 0
                        strOS = "Windows 2000" '1999年12月发布
                        If .wProductType = VER_NT_WORKSTATION Then
                            strOS = strOS & " " & "Professional"
                        Else
                            If bExInfo Then
                                If .wSuiteMask = VER_SUITE_ENTERPRISE Then
                                    strOS = strOS & " " & "Advanced Server"
                                ElseIf .wSuiteMask = VER_SUITE_DATACENTER Then
                                    strOS = strOS & " " & "Datacenter Server"
                                Else
                                    strOS = strOS & " " & "Server"
                                End If
                            End If
                        End If
                    Case 1
                        strOS = "Windows XP" '2001年8月发布
                        If .wSuiteMask = VER_SUITE_EMBEDDEDNT Then
                            strOS = strOS & " " & "Embedded"
                        ElseIf .wSuiteMask = VER_SUITE_PERSONAL Then
                            strOS = strOS & " " & "Home Edition"
                        Else
                            strOS = strOS & " " & "Professional"
                        End If
                    Case 2
                        If .wProductType = VER_NT_WORKSTATION And sysinfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64 Then
                            strOS = "Windows XP Professional x64 Edition"
                        ElseIf GetSystemMetrics(SM_SERVERR2) = 0 Then
                            strOS = "Windows Server 2003" '2003年3月发布
                        Else
                            strOS = "Windows Server 2003 R2"
                        End If
                        
                        If GetSystemMetrics(SM_SERVERR2) = 0 Then
                            If .wSuiteMask = VER_SUITE_BLADE Then
                                strOS = strOS & " " & "Web Edition"
                            ElseIf .wSuiteMask = VER_SUITE_COMPUTE_SERVER Then
                                strOS = strOS & " " & "Compute Cluster Edition"
                            ElseIf .wSuiteMask = VER_SUITE_STORAGE_SERVER Then
                                strOS = strOS & " " & "Storage Server"
                            ElseIf .wSuiteMask = VER_SUITE_DATACENTER Then
                                strOS = strOS & " " & "Datacenter Edition"
                            ElseIf .wSuiteMask = VER_SUITE_ENTERPRISE Then
                                strOS = strOS & " " & "Enterprise Edition"
                            End If
                        ElseIf .wSuiteMask = VER_SUITE_STORAGE_SERVER Then
                            strOS = strOS & " " & "Storage Server"
                        End If
                End Select
            Case 6
                Select Case .dwMinorVersion
                    Case 0
                        If .wProductType = VER_NT_WORKSTATION Then
                            strOS = "Microsoft Windows Vista"
                            If .wSuiteMask = VER_SUITE_PERSONAL Then
                                strOS = strOS & " " & "Home"
                            End If
                        Else
                            strOS = "Microsoft Windows Server 2008"
                            If .wSuiteMask = VER_SUITE_DATACENTER Then
                                strOS = strOS & " " & "Datacenter Server"
                            ElseIf .wSuiteMask = VER_SUITE_ENTERPRISE Then
                                strOS = strOS & " " & "Enterprise"
                            End If
                        End If
                    Case 1
                        If .wProductType = VER_NT_WORKSTATION Then
                            strOS = "Windows 7"
                        Else
                            strOS = "Windows Server 2008 R2"
                        End If
                    Case 2
                        If .wProductType = VER_NT_WORKSTATION Then
                            strOS = "Windows 8"
                        Else
                            strOS = "Windows Server 2012"
                        End If
                    Case 3
                        If .wProductType = VER_NT_WORKSTATION Then
                            strOS = "Windows 8.1"
                        Else
                            strOS = "Windows Server 2012 R2"
                        End If
                End Select
        End Select
    End With
    GetVersionInfo = strOS
End Function

Private Function CheckRepeatLogin(ByVal strIpAddress As String) As Boolean
    '检查是否有重复登录
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strProgram As String
    On Error GoTo Errhand
    
    strProgram = GetCallEXE
    strSQL = "Select A.UserName, A.Program, B.IP" & vbNewLine & _
            "From gv$Session A, zlClients B" & vbNewLine & _
            "Where A.Terminal = B.工作站" & vbNewLine & _
            "      And A.Terminal = (Select Terminal From v$Session Where AudsID = Userenv('SessionID') and RowNum =1)" & vbNewLine & _
            "      And A.Program =[1] And A.AudsID <> Userenv('SessionID')" & vbNewLine & _
            "      And B.IP <> [2]"

    Set rsTemp = OpenSQLRecord(strSQL, "检查重复工作站", strProgram, strIpAddress)
    If rsTemp.RecordCount = 0 Then '可以登录
        CheckRepeatLogin = False
        Exit Function
    Else
        MsgBox "局域网中存在相同名称的计算机登录," & vbCrLf & "对方IP是:[" & NVL(rsTemp!IP) & "]", vbInformation, gstrSysName
        CheckRepeatLogin = True
        Exit Function
    End If
    Exit Function
Errhand:
    MsgBox "检查同名计算机出错：" & Err.Description & ",请联系技术人员进行解决！", vbInformation, gstrSysName
End Function

Private Function GetCallEXE() As String
'功能：获取调用当前DLL的EXE名称
    Dim strPName As String, strFileName As String

    strPName = String(256, Chr(0))
    Call GetModuleFileName(0, strPName, 256)
    strFileName = Left(strPName, InStr(strPName, Chr(0)) - 1)
    strFileName = UCase(Mid(strFileName, InStrRev(strFileName, "\") + 1))
    GetCallEXE = strFileName
End Function

Private Function GetLISStation() As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能   得到独立新版LIS的站点
'返回   得到站点和站点名称  空为没有站点
'        有的组织方式为 ,1,2;,站点1,站点2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim str站点  As String, str站点名称 As String
    
    On Error GoTo Errhand
    '判断是否独立安装
    strSQL = "select 1 计数 from zlsystems where 编号 = 2500 and 共享号 is null"
    Set rsTmp = OpenSQLRecord(strSQL, "检查是否独立安装新版LIS")
    If rsTmp.EOF Then Exit Function
    '查找是否有默认的站点
    strSQL = "Select Distinct A.站点, B.名称" & vbNewLine & _
            "From (Select Distinct A.站点" & vbNewLine & _
            "       From 检验仪器记录 A, 检验仪器人员 B, 人员表 C,上机人员表 d" & vbNewLine & _
            "       Where A.Id = B.仪器id And A.站点 Is Not Null And B.人员id = C.Id and c.id = d.人员ID And d.用户名 = [1]) A, Zlnodelist B" & vbNewLine & _
            "Where A.站点 = B.编号" & vbNewLine & _
            "Order By A.站点"
    Set rsTmp = OpenSQLRecord(strSQL, "站点查询", gobjRelogin.DBUser)
    Do While Not rsTmp.EOF
        str站点 = str站点 & "," & rsTmp!站点
        str站点名称 = str站点名称 & "," & rsTmp!名称
        rsTmp.MoveNext
    Loop
    If str站点 <> "" Then
        GetLISStation = str站点 & ";" & str站点名称
    End If
    Exit Function
Errhand:
    MsgBox "获取LIS工作站出错：" & Err.Description & ",请联系技术人员进行解决！", vbInformation, gstrSysName
End Function

Private Sub UpdateEmrInterface()
    Dim objEMR As Object
    
    On Error Resume Next
    Err.Clear
    Set objEMR = CreateObject("zl9EmrInterface.ClsEmrInterface")
    If Err.Number = 0 Then
        Call objEMR.CheckUpdate1(gobjRelogin.InputUser, IIf(gobjRelogin.IsTransPwd, "", "[DBPASSWORD]") & gobjRelogin.InputPwd, IIf(InStr(UCase(gstrCommand), "ZLDOCUPDATER=1") > 0, False, True))
        If Err.Number <> 0 Then
            Err.Clear
            Call objEMR.CheckUpdate(gobjRelogin.InputUser, IIf(gobjRelogin.IsTransPwd, "", "[DBPASSWORD]") & gobjRelogin.InputPwd)
        End If
        Set gobjRelogin.EMR = objEMR
    Else
        Set gobjRelogin.EMR = Nothing
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Public Function UpdateZLHIS(ByVal strComputerName As String, Optional ByVal blnBrwCall As Boolean, Optional ByVal blnForceUpdate As Boolean) As Boolean
'功能：调用ZLHIS进行升级
'      blnBrwCall=是否导航台调用,导航台调用升级时检查预升级时点
    Dim strUpdateExe As String, strUpdateExePath As String
    Dim objFSO As New FileSystemObject
    Dim objConn As clsConnect, datCur           As Date
    Dim rsTemp As ADODB.Recordset, strSQL       As String
    Dim strJobName As String, blnDownload       As Boolean
    Dim strTmpPath As String, lngWait           As Long
    Dim strTmpGet  As String, blnMustNowUpdate  As Boolean
    
    strUpdateExe = "zlHisCrust.exe"
    gstrSetupPath = App.Path
    Call SaveSetting("ZLSOFT", "公共全局", "升级程序", UCase(strUpdateExe)) '用于ZLRegister中特殊判断
    '非授权的程序不自动升级
    strTmpGet = IIf(gobjRelogin.IsTransPwd, TranPasswd(gobjRelogin.InputPwd), gobjRelogin.InputPwd)
    If strTmpGet Like "未授权的程序:*" Then
        UpdateZLHIS = True
        Exit Function
    End If
    '没有服务器配置或文件清单，则不升级
    If Not IsHaveClientUpgradeSet(blnForceUpdate) Then '客户端修复时，进行消息提示。
        UpdateZLHIS = True
        Exit Function
    End If
    '没有升级，收集等任务，则自动退出升级
    If Not CheckJobs(strComputerName, strJobName, blnBrwCall, blnForceUpdate, blnMustNowUpdate) Then
        If blnForceUpdate Then
            MsgBox "当前只能进行预升级，无法进行客户端修复！", vbInformation, gstrSysName
        Else
            UpdateZLHIS = True
        End If
        Exit Function
    End If
    
    If strJobName = "OfficialUpgrade" And blnBrwCall Then
        If blnMustNowUpdate Then
            MsgBox "检测到系统需要进行重要的更新，1分钟后会进行升级，请及时保存正在书写的内容。", vbInformation, gstrSysName
            lngWait = 1 '设置升级等待时间
        Else
            If MsgBox("检测到系统需要升级，是否立即升级?" & vbNewLine & "选择否后请重新登录进行升级。", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                UpdateZLHIS = True
                Exit Function
            End If
        End If
    End If
    If IsDesinMode Then
        strUpdateExePath = "C:\APPSOFT\zlHisCrust.exe"
        strTmpPath = "C:\APPSOFT\ZLUPTMP"
    Else
        strUpdateExePath = gstrSetupPath & "\zlHisCrust.exe"
        strTmpPath = gstrSetupPath & "\ZLUPTMP"
    End If
    '升级程序不存在，则准备下载
    If Not objFSO.FileExists(strUpdateExePath) Then
        '先准备临时升级目录
        If Not objFSO.FolderExists(strTmpPath) Then
            objFSO.CreateFolder (strTmpPath)
        End If
        strTmpPath = strTmpPath & "\" & Format(Now, "YYMMDDHHmmss")
        If Not objFSO.FolderExists(strTmpPath) Then
            Call objFSO.CreateFolder(strTmpPath)
        End If
        strTmpPath = strTmpPath & "\zlHisCrust.exe"
        Set objConn = New clsConnect
        If Not objConn.GetFileConnect(strComputerName) Then
            MsgBox "无法连接客户端升级服务器""" & objConn.ServerPath & """,请联系管理员。", vbExclamation, gstrSysName
            Exit Function
        End If
        blnDownload = objConn.DownloadFile("ZLHISCRUST.EXE", strTmpPath)
        If blnDownload Then
            On Error Resume Next
            '先清理本地文件
            If objFSO.FileExists(strUpdateExePath) Then
                If FileSystem.GetAttr(strUpdateExePath) <> vbNormal Then
                     Call FileSystem.SetAttr(strUpdateExePath, vbNormal)
                End If
                Call objFSO.DeleteFile(strUpdateExePath)
            End If
            If Err.Number <> 0 Then Err.Clear
            '先复制到APPSOFT下，如果失败，则复制到APPLY下
            objFSO.CopyFile strTmpPath, strUpdateExePath, True
            If Err.Number <> 0 Then
                Err.Clear
                If IsDesinMode Then
                    strUpdateExePath = "C:\APPSOFT\APPLY\zlHisCrust.exe"
                Else
                    strUpdateExePath = gstrSetupPath & "\APPLY\zlHisCrust.exe"
                End If
                '先清理本地文件
                If objFSO.FileExists(strUpdateExePath) Then
                    If FileSystem.GetAttr(strUpdateExePath) <> vbNormal Then
                         Call FileSystem.SetAttr(strUpdateExePath, vbNormal)
                    End If
                    Call objFSO.DeleteFile(strUpdateExePath)
                End If
                If Err.Number <> 0 Then Err.Clear
                objFSO.CopyFile strTmpPath, strUpdateExePath, True
                If Err.Number <> 0 Then
                    Err.Clear
                    '是否是新版自动升级外壳，是的话，则可以直接从临时目录启动。
                    If UCase(GetFileDesInfo(strTmpPath, "ProductName")) = "ZLHISINSTALLUPDATE" Then
                        strUpdateExePath = strTmpPath
                    End If
                End If
            End If
        End If
        If strTmpPath <> strUpdateExePath Then
            On Error Resume Next
            '临时路径
            If objFSO.FileExists(strTmpPath) Then
                If FileSystem.GetAttr(strTmpPath) <> vbNormal Then
                     Call FileSystem.SetAttr(strTmpPath, vbNormal)
                End If
                Call objFSO.DeleteFile(strTmpPath)
            End If
            Call objFSO.DeleteFolder(objFSO.GetParentFolderName(strTmpPath))
        End If
        If Not objFSO.FileExists(strUpdateExePath) Then
            MsgBox "没有找到客户端自动升级工具" & strUpdateExe & "并且无法通过升级服务器下载，请与系统管理员联系。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    '预升级可以在导航台运行中进行
    If StartHisCrust(strUpdateExePath, strJobName, lngWait, strTmpGet) And strJobName <> "PreUpgrade" And strJobName <> "CheckUpgrade" Then
        Exit Function
    End If
    UpdateZLHIS = True
End Function

Private Function GetDefaultFileServer() As Integer
'功能：获取默认服务器
'返回：若没有服务器设置返回-1，存在，则任意返回一个服务器编号
    Dim intDefaultSever As Integer, intServerType   As Integer
    Dim blnReadOld      As Boolean
    Dim strSQL          As String, rsTmp            As ADODB.Recordset
    
    On Error Resume Next
    intDefaultSever = -1
    strSQL = "Select 编号 From Zltools.Zlupgradeserver Where 是否升级 = 1"
    Set rsTmp = OpenSQLRecord(strSQL, "获取升级服务器")
    If Err.Number <> 0 Then '可能管理员使用的管理工具与各个客户端不匹配
        Err.Clear
        blnReadOld = True
    ElseIf rsTmp.EOF Then
        blnReadOld = True
    End If
    On Error GoTo errH
    If Not blnReadOld Then
        intDefaultSever = Val(rsTmp!编号 & "")
    Else
        strSQL = "select 内容 from zlreginfo where 项目='升级类型'"
        Set rsTmp = OpenSQLRecord(strSQL, "检查使用的升级类型")
        If Not rsTmp.EOF Then
            intServerType = Val(rsTmp!内容 & "")
        End If
        If intServerType = 0 Then
            strSQL = "select replace(项目,'服务器目录','') as 服务器 from zlreginfo where 项目 like '服务器目录%' and 内容 is not null"
            Set rsTmp = OpenSQLRecord(strSQL, "检查是否存配置的文件共享服务器")
        Else
            strSQL = "select replace(项目,'FTP服务器','') as 服务器 from zlreginfo where 项目 like 'FTP服务器%' and 内容 is not null"
            Set rsTmp = OpenSQLRecord(strSQL, "检查是否存配置的FTP服务器")
        End If
        If Not rsTmp.EOF Then
            intDefaultSever = Val(rsTmp!服务器 & "")
        End If
    End If
    GetDefaultFileServer = intDefaultSever
    Exit Function
errH:
    GetDefaultFileServer = intDefaultSever
    If gblnTimer Then
        If ErrCenter() = 1 Then
            Resume
        End If
    Else
        MsgBox "获取缺省服务器出错：" & Err.Description, vbInformation, gstrSysName
        Err.Clear
    End If
End Function

Private Function IsHaveClientUpgradeSet(Optional ByVal blnMsg As Boolean) As Boolean
'功能：是否存在升级相关的配置。
'参数：blnMsg=结果为False的时候是否提示
'返回：IsHaveClientUpgradeSet=True:存在可升级文件与服务器配置，False-两者中至少一个缺失
    Dim intServerID As Integer
    Dim strSQL          As String, rsTmp            As ADODB.Recordset
    
    On Error GoTo errH
    IsHaveClientUpgradeSet = True
    '先判断是否存在可升级文件
    strSQL = "Select 1 可升级文件 From Zltools.Zlfilesupgrade Where Md5 Is Not Null And Rownum < 2"
    Set rsTmp = OpenSQLRecord(strSQL, "检查是否存在可升级文件")
    If Not rsTmp.EOF Then '可升级文件存在则需要进一步判定是否设置了升级服务器
        intServerID = GetDefaultFileServer
        If intServerID = -1 Then
            If blnMsg Then
                MsgBox "没有设置客户端升级文件服务器，无法进行客户端修复！", vbInformation, gstrSysName
            End If
            IsHaveClientUpgradeSet = False
        End If
    Else
        If blnMsg Then
            MsgBox "尚未配置升级文件清单，无法进行客户端修复！请联系管理员！", vbInformation, gstrSysName
        End If
        IsHaveClientUpgradeSet = False
    End If
    Exit Function
errH:
    IsHaveClientUpgradeSet = False
    If gblnTimer Then
        If ErrCenter() = 1 Then
            Resume
        End If
    Else
        MsgBox "检查升级文件清单出错：" & Err.Description, vbInformation, gstrSysName
        Err.Clear
    End If
End Function

Private Function CheckJobs(ByVal strComputerName As String, ByRef strJobName As String, Optional ByVal blnBrwCall As Boolean, Optional ByVal blnForceUpdate As Boolean, Optional ByRef blnMustNowUpdate As Boolean) As Boolean
'功能:检查并获取升级程序的任务
'      blnBrwCall=是否导航台调用,导航台调用升级时检查预升级时点
'      blnForceUpdate=导航台点击客户端修复时该参数为True
'      blnMustNowUpdate=是否现在必须升级
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim datCur As Date, blnOnlyOfficialUp As Boolean, blnOnlyPreUp As Boolean
    Dim blnPreUp As Boolean, blnOfficialUp As Boolean, blnPreComplete As Boolean, blnCollect As Boolean
    Dim strStartTime As String, strEndTime As String
    
    On Error GoTo errH
    strJobName = "": blnMustNowUpdate = False
    '以下代码一般不可能出错
    datCur = Currentdate
    '判断任务是否合理，获取是否启用了定时升级
    strSQL = "Select Max(内容) 内容 From zlRegInfo Where 项目='客户端升级日期'"
    Set rsTmp = OpenSQLRecord(strSQL, "检查定时升级")
    If rsTmp!内容 & "" <> "" Then
        If CDate(Format(datCur, "yyyy-MM-dd HH:mm:ss")) >= CDate(Format(NVL(rsTmp!内容), "yyyy-MM-dd HH:mm:ss")) Then
            blnOnlyOfficialUp = True '只能正式升级
        Else
            blnOnlyPreUp = True '只能预升级
        End If
    Else
        blnOnlyOfficialUp = True
    End If
    On Error Resume Next
    Set rsTmp = Nothing
    '可能没有是否预升级字段(因为预升级时候，数据库还没升级），因此需要错误忽略
    strSQL = "Select 预升时点,Nvl(是否预升级,0) 是否预升级, Nvl(预升完成, 0) 预升完成, Nvl(升级标志, 0) 升级标志, Nvl(收集标志, 0) 收集标志,Nvl(是否立即升级,0) 是否立即升级 From Zlclients Where 工作站 = [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "检查当前任务", strComputerName)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo errH
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            blnPreUp = rsTmp!是否预升级 = 1
            blnOfficialUp = rsTmp!升级标志 = 1
            blnPreComplete = rsTmp!预升完成 = 1
            blnCollect = rsTmp!收集标志 = 1
            strStartTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!预升时点, "HH:00:00")
            strEndTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!预升时点, "HH:59:59")
            blnMustNowUpdate = rsTmp!是否立即升级 = 1
        End If
    Else
        '优先新方式读取，失败再使用老方式，增加兼容性
        strSQL = "Select 预升时点,Nvl(预升完成, 0) 预升完成, Nvl(升级标志, 0) 升级标志, Nvl(收集标志, 0) 收集标志 From Zlclients Where 工作站 = [1]"
        Set rsTmp = OpenSQLRecord(strSQL, "检查当前任务", strComputerName)
        If Not rsTmp.EOF Then
            blnPreUp = rsTmp!升级标志 = 1
            blnOfficialUp = rsTmp!升级标志 = 1
            blnPreComplete = rsTmp!预升完成 = 1
            blnCollect = rsTmp!收集标志 = 1
            strStartTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!预升时点, "HH:00:00")
            strEndTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!预升时点, "HH:59:59")
        End If
    End If
    '当前只能进行预升级
    If blnOnlyPreUp Then
        '有预升级任务
        If blnPreUp Or blnOfficialUp Then
            If Not blnPreComplete Then
                If datCur >= CDate(strStartTime) And datCur <= CDate(strEndTime) Then
                    strJobName = "PreUpgrade"
                Else
                    Exit Function
                End If
            Else
                Exit Function
            End If
        '没有预升级任务，但是有收集任务
        ElseIf blnCollect Then
            strJobName = "CheckUpgrade"
        Else
            Exit Function
        End If
    '当前只能进行正式升级
    ElseIf blnOnlyOfficialUp Then
        If blnForceUpdate Then
            strJobName = "Repair"
        Else
            '有正式升级任务
            If blnOfficialUp Then
                strJobName = "OfficialUpgrade"
            '没有正式升级任务，但是有收集任务
            ElseIf blnCollect Then
                strJobName = "CheckUpgrade"
            Else
                Exit Function
            End If
        End If
    End If
    CheckJobs = True
    Exit Function
errH:
    If gblnTimer Then
        If ErrCenter() = 1 Then
            Resume
        End If
    Else
        MsgBox "检查客户端任务出错：" & Err.Description, vbInformation, gstrSysName
        Err.Clear
    End If
End Function

Public Function DeCipher(ByVal strText As String) As String
'密码解密程序
    Const MIN_ASC = 32    '最小ASCII码
    Const MAX_ASC = 126 '最大ASCII码 字符
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intlen As Integer, intSeedLen As Integer
    Dim intStart As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    
    If strText = "" Then Exit Function
    '随机种子长度
    intSeedLen = Asc(Mid(strText, 1, 1)) - MIN_ASC
    intlen = Len(strText)
    '采用旧的随机算法
    If intSeedLen > 0 And intSeedLen < intlen - 3 And intSeedLen < 5 Then
        '获取随机种子
        '随机种子的随机数为999
        Rnd (-1)
        Randomize (999)
        For i = 2 To 1 + intSeedLen
            intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
            If intChr >= MIN_ASC And intChr <= MAX_ASC Then
                intChr = intChr - MIN_ASC
                lngOffset = Int((NUM_ASC + 1) * Rnd())
                intChr = ((intChr - lngOffset) Mod NUM_ASC)
                If intChr < 0 Then
                    intChr = intChr + NUM_ASC
                End If
                intChr = intChr + MIN_ASC
                strDeText = strDeText & Chr(intChr)
            End If
        Next
        If Not IsNumeric(strDeText) Then
            strDeText = "123"
            intStart = 1
        Else
            intStart = 2 + intSeedLen
        End If
    Else
        strDeText = "123"
        intStart = 1
    End If
        
    '内容解密的种子
    Rnd (-1)
    Randomize (Val(strDeText))
    strDeText = ""
    For i = intStart To intlen
        intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr - lngOffset) Mod NUM_ASC)
            If intChr < 0 Then
                intChr = intChr + NUM_ASC
            End If
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        End If
    Next
    DeCipher = strDeText
End Function

Public Function GetLastDllErr(Optional ByVal lngErr As Long) As String
    Dim strReturn As String
    If lngErr = 0 Then
        lngErr = GetLastError
    End If
    If lngErr = ERROR_EXTENDED_ERROR Then
        GetLastDllErr = GetWNetErr(lngErr)
    Else
        strReturn = String$(256, 32)
        FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lngErr, 0&, strReturn, Len(strReturn), ByVal 0
        strReturn = Trim(strReturn)
        GetLastDllErr = Replace(Replace(strReturn, Chr(10), ""), Chr(13), "")
    End If
End Function

Private Function GetWNetErr(ByVal lngErr As Long) As String
    Dim strErr As String * 256
    Dim strName As String * 256
    Dim lngRet As Long
    lngRet = WNetGetLastError(lngErr, strErr, Len(strErr), strName, Len(strName))
    GetWNetErr = Replace(Replace("[" & TruncZero(strName) & "]" & TruncZero(strErr), Chr(10), ""), Chr(13), "")
End Function

Public Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Private Function GetFileDesInfo(ByVal strFileName As String, ByVal strEntryName As String) As String
    Dim i               As Long
    Dim lngVerSize      As Long
    Dim bytVerBlock()   As Byte
    Dim strSubBlock  As String
    Dim bytTranslate()  As Byte, lngAdrTranslate    As Long, lngTranslateSize       As Long
    Dim bytBuffer()     As Byte, lngBuffer          As Long, lngAdrBuffer           As Long

    On Error GoTo errH
    lngVerSize = GetFileVersionInfoSize(strFileName, 0&)
    If lngVerSize <= 0 Then Exit Function
    
    ReDim bytVerBlock(lngVerSize - 1)
    Call GetFileVersionInfo(strFileName, 0&, lngVerSize, bytVerBlock(0))
    
    VerQueryValue VarPtr(bytVerBlock(0)), "\\VarFileInfo\\Translation", lngAdrTranslate, lngTranslateSize
    ReDim bytTranslate(lngTranslateSize - 1)
    CopyMemory bytTranslate(0), ByVal lngAdrTranslate, lngTranslateSize
    For i = 1 To lngTranslateSize / (UBound(bytTranslate) + 1)
        strSubBlock = "\\StringFileInfo\\"
        strSubBlock = strSubBlock & Byte2Hex(bytTranslate(), 0, 1, True)
        strSubBlock = strSubBlock & Byte2Hex(bytTranslate(), 2, 3, True)
        strSubBlock = strSubBlock & "\\" & strEntryName
        
        VerQueryValue VarPtr(bytVerBlock(0)), strSubBlock, lngAdrBuffer, lngBuffer
        If lngAdrBuffer <> 0 And lngBuffer <> 0 Then
            ReDim bytBuffer(lngBuffer - 1)
            CopyMemory bytBuffer(0), ByVal lngAdrBuffer, lngBuffer
            ReDim Preserve bytBuffer(InStrB(bytBuffer, ChrB(0)) - 2)
            GetFileDesInfo = StrConv(bytBuffer, vbUnicode)
        End If
    Next
    Exit Function
errH:
    Err.Clear
End Function
Private Function Byte2Hex(bytArray() As Byte, Optional ByVal lngStart As Long = 0, Optional ByVal lngEnd As Long = -1, Optional fReversed As Boolean = False) As String
    Dim i     As Long
    lngStart = IIf(lngStart < 0, 0, lngStart)
    lngEnd = IIf(lngEnd < 0, UBound(bytArray), lngEnd)
    
    If fReversed Then
        For i = lngEnd To lngStart Step -1
            Byte2Hex = Byte2Hex & Right$("00" & Hex(bytArray(i)), 2)
        Next
    Else
        For i = lngStart To lngEnd
            Byte2Hex = Byte2Hex & Right$("00" & Hex(bytArray(i)), 2)
        Next
    End If
End Function

Public Function ComputerName() As String
    '******************************************************************************************************************
    '功能：获取电脑名称
    '参数：
    '说明：
    '******************************************************************************************************************
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    ComputerName = strComputer
    ComputerName = Trim(Replace(ComputerName, Chr(0), ""))
End Function

Public Function LoacalIP(Optional ByRef strErr As String) As String
    '功能：通过API获取临时IP
    
    Dim ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim TempList() As String
    Dim TempIP As String
    Dim Tempi As Long
    Dim Listing As MIB_IPADDRTABLE
    Dim L3 As String
    Dim strTmpErr As String, strALLErr As String
    
    strErr = ""
    On Error GoTo Errhand
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
        TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr, strTmpErr)
        If strTmpErr <> "" Then strALLErr = strALLErr & IIf(strALLErr = "", "", "|") & strTmpErr
    Next Tel
    'Sort Out The IP For WAN
        TempIP = TempList(0)
        For Tempi = 0 To Listing.dEntrys - 1
            L3 = Left(TempList(Tempi), 3)
            If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
                TempIP = TempList(Tempi)
            End If
        Next Tempi
        LoacalIP = TempIP 'Return The TempIP
    Exit Function
    strErr = strALLErr
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    strErr = strALLErr & IIf(strALLErr = "", "", "|") & Err.Description
    Err.Clear
End Function

Private Function ConvertAddressToString(longAddr As Long, Optional ByRef strErr As String) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    
    strErr = ""
    On Error GoTo errH
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errH:
    strErr = Err.Description
    Err.Clear
End Function

Public Function IsDesinMode() As Boolean
'功能： 确定当前模式为设计模式
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
End Function

Public Function FromatSQL(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'参数：strText=处理字符
'         blnCrlf=是否去掉换行符
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    FromatSQL = strText
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    Dim arrPars() As Variant, i As Long
    arrPars = arrInput
    If gblnTimer Then
        Set OpenSQLRecord = zlDatabase.OpenSQLRecordByArray(strSQL, strTitle, arrPars)
    Else
        Set OpenSQLRecord = OpenSQLRecordByArray(strSQL, strTitle, arrPars)
    End If
End Function

Private Function OpenSQLRecordByArray(ByVal strSQL As String, ByVal strTitle As String, arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'               因为使用绑定变量,对带"'"的字符参数,不需要使用"''"形式。
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'      cnOracle=当不使用公共连接时传入
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    Dim strSQLTmp As String, arrstr As Variant
    Dim strTmp As String, strSQLtmp1 As String
    Dim lngErrNum As Long, strErrInfo As String
    
    '检查如果使用了动态内存表，并且没有使用/*+ XXX*/等提示字时自动加上
    strSQLTmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLTmp, 7)), 1, 2) <> "/*" And Mid(strSQLTmp, 1, 6) = "SELECT" Then
        arrstr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrstr)
            strSQLtmp1 = strSQLTmp
            Do While InStr(strSQLtmp1, arrstr(i)) > 0
                '判断前面是否用了IN 用了则不加Rule
                '先找到最近一个SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrstr(i)) - 1)
                strTmp = Replace(FromatSQL(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  '取后面3个字符
                
                If strTmp = "IN(" Then '属于in(select这种情况，则继续?罚词欠翊嬖诿挥惺褂谜庵中捶ǖ钠渌诖婧?
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrstr(i)) + Len(arrstr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrstr) Then
            strSQL = "Select /*+ RULE*/" & Mid(Trim(strSQL), 7)
        End If
    End If
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop
    
    If UBound(arrInput) + 1 < intMax Then
        Err.Raise 9527, strTitle, "SQL语句绑定变量不全，调用来源：" & strTitle
    End If

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '数组
            '这种方式可用于一些IN子句或Union语句
            '表示同一个参数的多个值,参数号不可与其它数组的参数号交叉,且要保证数组的值个数够用
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '字符
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '日期
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '该参数在数组中用到第几个值了
        End Select
    Next
'    If gblnSys = True Then
'        Set cmdData.ActiveConnection = gcnSysConn
'    Else
    Set cmdData.ActiveConnection = gcnOracle '这句比较慢(这句执行1000次约0.5x秒)
'    End If
    cmdData.CommandText = strSQL
    
'    Call gobjComLib.SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecordByArray = cmdData.Execute
    Set OpenSQLRecordByArray.ActiveConnection = Nothing
'    Call gobjComLib.SQLTest
End Function

Public Sub ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'功能：执行过程语句,并自动对过程参数进行绑定变量处理
'参数：strSQL=过程语句,可能带参数,形如"过程名(参数1,参数2,...)"。
'      cnOracle=当不使用公共连接时传入
'说明：以下几种情况过程参数不使用绑定变量,仍用老的调用方法：
'  1.参数部份是表达式,这时程序无法处理绑定变量类型和值,如"过程名(参数1,100.12*0.15,...)"
'  2.中间没有传入明确的可选参数,这时程序无法处理绑定变量类型和值,如"过程名(参数1, , ,参数3,...)"
'  3.因为该过程是自动处理,不是一定使用绑定变量,对带"'"的字符参数,仍要使用"''"形式。
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    Dim lngErrNum As Long, strErrInfo As String
    
    If Right(Trim(strSQL), 1) = ")" Then
        '执行的过程名
        strTemp = Trim(strSQL)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        '执行过程参数
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For i = 1 To Len(strTemp)
            '是否在字符串内，以及表达式的括号内
            If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                strPar = Trim(strPar)
                With cmdData
                    If IsNumeric(strPar) Then '数字
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, strPar)
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '字符串
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        
                        'Oracle连接符运算:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '双"''"的绑定变量处理
                        If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                        
                        '电子病历处理LOB时，如果用绑定变量转换为RAW时超过2000个字符要用adLongVarChar
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
                        If intMax <= 2000 Then
                            intMax = IIf(intMax <= 200, 200, 2000)
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMax, strPar)
                        Else
                            If intMax < 4000 Then intMax = 4000
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adLongVarChar, adParamInput, intMax, strPar)
                        End If
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '日期
                        strPar = Split(strPar, "(")(1)
                        strPar = Trim(Split(strPar, ",")(0))
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        If strPar = "" Then
                            'NULL值当成数字处理可兼容其他类型
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(strPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , CDate(strPar))
                        End If
                    ElseIf UCase(strPar) = "SYSDATE" Then '日期
                        If datCur = CDate(0) Then datCur = Currentdate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then 'NULL值当成字符处理可兼容其他类型
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, 200, Null)
                    ElseIf strPar = "" Then '可选参数当成NULL处理可能改变了缺省值:因此可选参数不能写在中间
                        GoTo NoneVarLine
                    Else '可能是其他复杂的表达式，无法处理
                        GoTo NoneVarLine
                    End If
                End With
                
                strPar = ""
            Else
                strPar = strPar & Mid(strTemp, i, 1)
            End If
        Next
        
        '程序员调用过程时书写错误
        If blnStr Or intBra <> 0 Then
            Err.Raise -2147483645, , "调用 Oracle 过程""" & strProc & """时，引号或括号书写不匹配。原始语句如下：" & vbCrLf & vbCrLf & strSQL
            Exit Sub
        End If
        
        '补充?号
        strTemp = ""
        For i = 1 To cmdData.Parameters.Count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢
        cmdData.CommandType = adCmdText
        cmdData.CommandText = strProc
        
'        Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
        Call cmdData.Execute
'        Call gobjComLib.SQLTest
    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:
'    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
    '说明：为了兼容新连接方式
    '1.新连接用adCmdStoredProc方式在8i下面有问题
    '2.新连接如果不使用{},则即使过程没有参数也要加()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    gcnOracle.Execute strSQL, , adCmdText
'    Call gobjComLib.SQLTest
End Sub


Public Function IP(Optional ByVal strErr As String) As String
    '******************************************************************************************************************
    '功能:通过oracle获取的计算机的IP地址
    '入参:strDefaultIp_Address-缺省IP地址
    '出参:
    '返回:返回IP地址
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim strIp_Address As String
    Dim strSQL As String
        
    On Error GoTo Errhand
    
    strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
    Set rsTmp = OpenSQLRecord(strSQL, "获取IP地址")
    If rsTmp.EOF = False Then
        strIp_Address = NVL(rsTmp!Ip_Address)
    End If
    If strIp_Address = "" Then strIp_Address = LoacalIP(strErr)
    If Replace(strIp_Address, " ", "") = "0.0.0.0" Then strIp_Address = ""
    IP = strIp_Address
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    strErr = strErr & IIf(strErr = "", "", "|") & Err.Description
    Err.Clear
End Function

Public Function Currentdate() As Date
    '-------------------------------------------------------------
    '功能：提取服务器上当前日期
    '参数：
    '返回：由于Oracle日期格式的问题，所以
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngErrNum As Long, strErrInfo As String
    
    Err = 0
    On Error GoTo errH
    With rsTemp
        .CursorLocation = adUseClient
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    Currentdate = rsTemp.fields(0).Value
    rsTemp.Close
    Exit Function
errH:
    Currentdate = 0
    Err = 0
End Function

Public Function CheckPWDComplex(ByRef cnInput As ADODB.Connection, ByVal strChcekPWD As String, Optional ByRef strToolTip As String) As Boolean
'功能：检查密码复杂度
'参数：cnInput=传入的连接
'          strChcekPWD=等待检查的密码
'          strToolTip=鼠标提示生成
'返回：True-检查成功；False-检查失败
    Dim strSQL As String, rsData As New ADODB.Recordset
    Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
    Dim blnPwdLen As Boolean, intPwdMin As Integer, intPwdMax As Integer
    Dim blnComplex As Boolean, strOterChrs As String
    Dim lngLen As Long, i As Integer, intChr As Integer
    
    On Error GoTo errH
    strToolTip = ""
    strSQL = "Select 参数号,Nvl(参数值,缺省值) 参数值 From zlOptions Where 参数号 in (20,21,22,23)"
    rsData.Open strSQL, cnInput
    blnPwdLen = False: intPwdMin = 0: intPwdMax = 0
    blnComplex = False: strOterChrs = ""
    Do While Not rsData.EOF
        Select Case rsData!参数号
            Case 20 '是否控制密码长度
                blnPwdLen = Val(rsData!参数值 & "") = 1
            Case 21 '密码长度下限
                intPwdMin = Val(rsData!参数值 & "")
            Case 22 '密码长度上限
                intPwdMax = Val(rsData!参数值 & "")
            Case 23 '是否控制密码复杂度
                blnComplex = Val(rsData!参数值 & "") = 1
        End Select
        rsData.MoveNext
    Loop
    '生成悬浮提示
    If blnPwdLen Then
        If intPwdMin = intPwdMax Then
            strToolTip = "密码必须为" & intPwdMax & " 位字符。"
        Else
            strToolTip = "密码必须为" & intPwdMin & "至" & intPwdMax & " 位字符。"
        End If
     End If
     If blnComplex Then
        If strToolTip <> "" Then
            strToolTip = strToolTip & vbNewLine & "至少包含一个数字、一个字母与一个特殊字符组成。"
        Else
            strToolTip = "至少由一个数字、一个字母与一个特殊字符组成。"
        End If
     End If
    '长度检查
    lngLen = ActualLen(strChcekPWD)
    If lngLen <> Len(strChcekPWD) Then
        MsgBox "新密码包含双字节字符，请检查！", vbInformation, gstrSysName
        Exit Function
    End If
    If blnPwdLen Then
        If Not (lngLen >= intPwdMin And lngLen <= intPwdMax) Then
            If intPwdMin = intPwdMax Then
                MsgBox "密码必须为" & intPwdMax & " 位字符！", vbInformation, gstrSysName
                Exit Function
            Else
                MsgBox "密码必须为" & intPwdMin & "至" & intPwdMax & " 位字符！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    For i = 1 To Len(strChcekPWD)
        intChr = Asc(UCase(Mid(strChcekPWD, i, 1)))
        If intChr >= 32 And intChr < 127 Then
            'Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
            Select Case intChr
                Case 48 To 57 '数字
                    blnHaveNum = True
                Case 65 To 90 '字母
                    blnAlpha = True
                Case 32, 34, 47, 64  '空格,双引号,/,@
                    strOterChrs = strOterChrs & Chr(intChr)
                Case Is < 48, 58 To 64, 91 To 96, Is > 122
                    blnChar = True
            End Select
        Else
            strOterChrs = strOterChrs & Chr(intChr)
        End If
    Next
    If strOterChrs <> "" Then
        MsgBox "密码不容许有以下字符：" & strOterChrs, vbInformation, gstrSysName
        Exit Function
    ElseIf Not (blnHaveNum And blnAlpha And blnChar) And blnComplex Then
        MsgBox "密码至少由一个数字、一个字母与一个特殊字符组成。", vbInformation, gstrSysName
        Exit Function
    End If
    CheckPWDComplex = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox Err.Description, vbInformation, gstrSysName
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '功能：求取指定字符串的实际长度，用于判断实际包含双字节字符串的
    '       实际数据存储长度
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function OpenImeByName(Optional strIme As String) As Boolean
'功能:按名称打开中文输入法,不指定名称时关闭中文输入法。支持部分名称。
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim blnNotCloseIme As Boolean
    
    If strIme = "不自动开启" Then OpenImeByName = True: Exit Function
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    blnNotCloseIme = True
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            ImmGetDescription arrIme(lngCount), strName, Len(strName)
            If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenImeByName = True
                Exit Function
            End If
        ElseIf strIme = "" Then
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenImeByName = True
            Exit Function
        End If
    Loop Until lngCount = 0
    
    If blnNotCloseIme And strIme = "" Then
        '由于windows Vista系统的英文输入法用ImmIsIME测试出是true的输入法,因此,需要单独处理.
        '刘兴宏:2008/09/03
        If ActivateKeyboardLayout(arrIme(0), 0) <> 0 Then OpenImeByName = True: Exit Function
    End If
End Function
