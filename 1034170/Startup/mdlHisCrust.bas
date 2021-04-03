Attribute VB_Name = "mdlHisCrust"
Option Explicit

'���������������API
'----------------------------------------------------------------------------------------------------
'Window�汾����
'win2000 ���°汾
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
'��ȡ�ڴ�
Private Type MEMORYSTATUS  'win2000�����°汾
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
'ȡӲ�̴�С
Private Const DRIVE_FIXED = 3
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const STRSPLIT As String = "���"

'API������Ϣ��ȡ
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
Private Const ERROR_EXTENDED_ERROR          As Long = 1208
'�ļ�������Ϣ�ж�
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (ByVal pBlock As Long, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
'Public Const FVN_Comments           As String = "Comments"          'ע��
'Public Const FVN_InternalName       As String = "InternalName"      '�ڲ�����
'Public Const FVN_ProductName        As String = "ProductName"       '��Ʒ��
'Public Const FVN_CompanyName        As String = "CompanyName"       '��˾��
'Public Const FVN_ProductVersion     As String = "ProductVersion"    '��Ʒ�汾
'Public Const FVN_FileDescription    As String = "FileDescription"   '�ļ�����
'Public Const FVN_OriginalFilename   As String = "OriginalFilename"  'ԭʼ�ļ���
'Public Const FVN_FileVersion        As String = "FileVersion"       '�ļ��汾
'Public Const FVN_SpecialBuild       As String = "SpecialBuild"      '��������
'Public Const FVN_PrivateBuild       As String = "PrivateBuild"      '˽�б����
'Public Const FVN_LegalCopyright     As String = "LegalCopyright"    '�Ϸ���Ȩ
'Public Const FVN_LegalTrademarks    As String = "LegalTrademarks"   '�Ϸ��̱�
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'hModule��һ��ģ��ľ����������һ��DLLģ�飬������һ��Ӧ�ó����ʵ�����������ò���ΪNULL���ú������ظ�Ӧ�ó���ȫ·��?
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
'��������(ComputerName)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'��ʱIP��ȡ
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

Public gstrExeFile      As String '���õ�¼������EXE·��
Public gstrSetupPath    As String 'APPSOFT·��
Public glnghInstance    As Long
Public gblnTimer            As Boolean  '�Ƿ�ʱ�������Ŀͻ��˸��¼��

Public Function CheckAllowByTerminal() As Boolean
'����:����Ƿ�����ʹ�ñ�����վ,�Լ����е�ǰ����վ��Ϣ�ĵǼ�
'     �ж��Ƿ�����ù���վʹ�ó���
'     �����Ҫ�滻���ز�������ִ���滻�����������Ҫ�������������ǳ��򣬲��ر��˳�
'����:�ɹ�,����true,���򷵻�False
'���棺���ڻ�û�г�ʼ���������������Ӷ��󣬸ú����в���ʹ�ù��������е����ݿ���ʷ���

    Dim rsTmp As ADODB.Recordset, strSQL As String, strRowID As String '�ͻ��˵�ROWID
    Dim strComuterInfo As String, arrComputer As Variant, strComputerName As String, strIpAddress As String
    Dim strTmp As String, arrTmp As Variant, i As Integer
    Dim bln���վ�� As Boolean, lng��վ�� As Long, bln��վ�� As Boolean, bln��վ�� As Boolean
    Dim strվ��       As String, strվ���� As String, str���� As String, strȱʡ����
    Dim blnAllow As Boolean, blnUpdate As Boolean
    Dim int��������� As Integer, int������ƵԴ As Integer, int������ As Integer, int������־ As Integer
    
'    Call SQLTest(App.EXEName, "mdlHisCrust", "�°���Ӳ����Զ��������")
    Call UpdateEmrInterface '�°���Ӳ����Զ�����
'    Call SQLTest

    strIpAddress = IP '��oracle���ӵ�IP��ַΪ��
    strComputerName = ComputerName
    '����Ƿ�����������
    If CheckRepeatLogin(strIpAddress) = True Then
        CheckAllowByTerminal = False
        Exit Function
    End If
    '�ж��Ƿ�����ʹ��
    strComuterInfo = AnalyseConfigure
    arrComputer = Split(strComuterInfo, STRSPLIT)
    '1.��վ�������
    If Err.Number <> 0 Then Err.Clear
    On Error Resume Next
    strSQL = "Select Rowid as ID,վ��,����,Nvl(��ֹʹ��,0) as ����,Nvl(������־,0) as ����,Nvl(�ռ���־,0) as �ռ�,������,������ƵԴ From zlClients Where ����վ=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "��鹤��վ-��վ��Ϊ��", strComputerName)
    '��������δ��Ȩ��ԭ�򣬵��²�ѯ������ʱ������ʾ��ֹ��¼
    If rsTmp Is Nothing Then
        MsgBox Err.Description & vbNewLine & "������������ϵͳ��������ϵϵͳ����Ա���½��н�ɫ��Ȩ��", vbInformation, gstrSysName
        Exit Function
    End If
    '2.δ���ִ�վ��,����IP��ʽ���ң���ֻ��һ��ʱ�Ÿ��¼�����
    If rsTmp.EOF Then
        strSQL = "Select Rowid as ID,վ��,����, Nvl(��ֹʹ��,0) as ����,Nvl(������־,0) as ����,Nvl(�ռ���־,0) as �ռ�,������,������ƵԴ From zlClients Where IP=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "��鹤��վ-��վ��Ϊ��", strIpAddress)
        If rsTmp.RecordCount > 1 Then
            '������������,���CPU,�ڴ�,Ӳ��Ϊ��������.
            strSQL = "" & _
                "   Select Rowid as ID,վ��,����,Nvl(��ֹʹ��,0) as ����,Nvl(������־,0) as ����,Nvl(�ռ���־,0) as �ռ�,������,������ƵԴ " & _
                "   From zlClients Where IP=[1] and CPU=[2] and  �ڴ�=[3] and Ӳ��=[4]"
            Set rsTmp = OpenSQLRecord(strSQL, "��鹤��վ-��վ��Ϊ��", strIpAddress, CStr(arrComputer(2)), CStr(arrComputer(3)), CStr(arrComputer(4)))
        End If
    End If
    bln���վ�� = True
    '��������ڶ��,����ܴ���IP��ͻ�����,��˲����ж���Ҫ������ص�վ��.ֻ�ܵ����µ�վ���ϴ�
    If rsTmp.RecordCount > 1 Or rsTmp.EOF Then
        strRowID = ""
    Else '��ʾ������ص���Ϣ
        strRowID = NVL(rsTmp!id)
        int������ƵԴ = Val(NVL(rsTmp!������ƵԴ))
        '�������½,�������û�ѡ��,ֱ�Ӷ�ȡ
        If gstrCommand <> "" Then
            '�·���
            If InStr(gstrCommand, "ZLHISCRUSTCALL=1") > 0 And InStr(gstrCommand, "USER=") > 0 And InStr(gstrCommand, "PASS=") > 0 Then
                bln���վ�� = False
                strվ���� = NVL(rsTmp!վ��)
                gobjRelogin.DeptName = NVL(rsTmp!����)
            '�ϵ��жϷ���
            ElseIf InStrRev(gstrCommand, "/", -1) > 0 And InStrRev(gstrCommand, ",", -1) = 0 Then
                bln���վ�� = False
                strվ���� = NVL(rsTmp!վ��)
                gobjRelogin.DeptName = NVL(rsTmp!����)
            End If
        End If
        blnAllow = Val(rsTmp!���� & "") = 0
        int������ = Val(rsTmp!������ & "")  '0-��ʾ������
        blnUpdate = Val(rsTmp!���� & "") = 1
        If Not blnUpdate Then blnUpdate = Val(rsTmp!�ռ� & "") = 1
    End If

    If bln���վ�� Then
        strSQL = "Select b.����, a.վ��, a.ȱʡ" & vbNewLine & _
                "From (Select c.վ��, b.ȱʡ" & vbNewLine & _
                "       From �ϻ���Ա�� a, ������Ա b, ���ű� c" & vbNewLine & _
                "       Where a.��Աid = b.��Աid And b.����id = c.Id And a.�û��� = Upper([1])) a, Zlnodelist b" & vbNewLine & _
                "Where a.վ�� = b.���(+)" & vbNewLine & _
                "Order By վ��"
        Set rsTmp = OpenSQLRecord(strSQL, "��鲢ȷ������Ժ��", gobjRelogin.DBUser)
        If rsTmp Is Nothing Then
            MsgBox Err.Description & vbNewLine & "������������ϵͳ��������ϵϵͳ����Ա���½��н�ɫ��Ȩ��", vbInformation, gstrSysName
            Exit Function
        End If
        Do While Not rsTmp.EOF
            If NVL(rsTmp!վ��, "") <> "" Then
                strվ�� = strվ�� & "," & NVL(rsTmp!վ��, "")
                str���� = str���� & "," & NVL(rsTmp!����)
                lng��վ�� = lng��վ�� + 1
            Else
                bln��վ�� = True
            End If
            If NVL(rsTmp!ȱʡ, "0") = 1 Then
                strȱʡ���� = NVL(rsTmp!����)
            End If
            rsTmp.MoveNext
        Loop
        '�����ǰ��¼��Ա�������Ŷ�û������վ�㣬���������ڲ��Ҹ�Ժ�Ƿ�������վ�����!
        If strվ�� = "" Or (bln��վ�� And lng��վ�� <> 1) Then
            '������װ�°�LISʱҲ��Ҫ��������ȡվ��
            strTmp = GetLISStation()
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ";")
                strվ�� = arrTmp(0)
                str���� = arrTmp(1)
            Else
                strվ�� = "": str���� = ""
                strSQL = "select distinct (A.վ��),B.���� from ���ű� A,zlNodeList B where A.վ��=B.��� And A.վ�� is not null order by A.վ��"
                Set rsTmp = OpenSQLRecord(strSQL, "����Ƿ�����վ�����")
                If Not rsTmp Is Nothing Then
                    Do While Not rsTmp.EOF
                        If NVL(rsTmp!վ��, "") <> "" Then
                            strվ�� = strվ�� & "," & NVL(rsTmp!վ��, "")
                            str���� = str���� & "," & NVL(rsTmp!����)
                        End If
                        rsTmp.MoveNext
                    Loop
                End If
            End If
        End If
        If strվ�� <> "" Then
            strվ�� = Mid(strվ��, 2)
            str���� = Mid(str����, 2)
            arrTmp = Split(strվ��, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                If i = LBound(arrTmp) Then
                    strվ���� = arrTmp(i)
                Else
                    If strվ���� <> arrTmp(i) Then
                        bln��վ�� = True
                        Exit For
                    End If
                End If
            Next
            If bln��վ�� Then '��ʾ�û�ѡ��ǰ�����λ�����ڵĲ��š�
                strվ���� = GetSetting("ZLSOFT", "˽��ģ��\" & gobjRelogin.DBUser & "\" & App.ProductName & "\" & App.EXEName, "��ǰվ��ѡ��", "")
                Call frmSelClient.ShowEdit(strվ��, str����, strվ����)
                strվ���� = IIf(frmSelClient.gstrվ�� = "��", "", frmSelClient.gstrվ��)
                gobjRelogin.DeptName = frmSelClient.gstrCurվ��
                Call SaveSetting("ZLSOFT", "˽��ģ��\" & gobjRelogin.DBUser & "\" & App.ProductName & "\" & App.EXEName, "��ǰվ��ѡ��", strվ����)
            End If
        End If
    End If
    gobjRelogin.NodeNo = IIf(strվ���� <> "", strվ����, "-")
    If gobjRelogin.DeptName = "" Then gobjRelogin.DeptName = strȱʡ����
    If strRowID = "" Then '�����Ĺ���վ����û�иù���վ�����ݣ��ϴ���IP����������CPU���ڴ桢Ӳ�̡�����ϵͳ��
        int��������� = GetDefaultFileServer
        If int��������� = -1 Then '��ȡĬ�Ϸ�����ʧ�ܣ����������ָ���������ŵĳ�ʼֵ
            int��������� = 0
            int������־ = 0
        Else
            int������־ = 1
        End If
        strSQL = "Zl_Zlclients_Set(0,Null,'" & strComputerName & "','" & strIpAddress & "','" & arrComputer(2) & "','" & arrComputer(3) & _
                    "','" & arrComputer(4) & "','" & arrComputer(5) & "','" & gobjRelogin.DeptName & "',Null,Null," & int��������� & "," & int������־ & _
                    ",0,'" & strվ���� & "',0,Null,Null," & int������ƵԴ & ")"
        ExecuteProcedure strSQL, "��������վ"
        '�����ͻ��˲���������ֱ���˳�
        If int������־ = 0 Then
            CheckAllowByTerminal = True
            Exit Function
        End If
        blnUpdate = True
    Else
        strSQL = "Zl_Zlclients_Set(1,'" & strRowID & "','" & strComputerName & "','" & strIpAddress & "','" & arrComputer(2) & "','" & arrComputer(3) & _
                    "','" & arrComputer(4) & "','" & arrComputer(5) & "','" & gobjRelogin.DeptName & "',Null,Null,Null,Null," & int������ & ",'" & strվ���� & "',0,Null,Null," & int������ƵԴ & ")"
        '��Ҫ������ص�վ����Ϣ
        ExecuteProcedure strSQL, "���¹���վ"
        If Not blnAllow Then
            MsgBox "�ù���վ�ѱ�����Ա���ã�", vbInformation, gstrSysName
            Exit Function
        End If
        '�������������
        If int������ > 0 Then
            strSQL = "Select SID From gv$Session Where Upper(PROGRAM) Like 'ZL%.EXE' And Status<>'KILLED' And MACHINE=(Select Max(MACHINE) From v$Session Where AUDSID=UserENV('SessionID'))"
            Set rsTmp = OpenSQLRecord(strSQL, "�����������")
            If rsTmp.RecordCount > int������ Then
                MsgBox "��ǰ����վ���ֻ���� " & int������ & " ����¼���ӣ���ǰ�Ѿ��� " & rsTmp.RecordCount - 1 & " �����ӡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    On Error GoTo Errhand
AutoUpGrude:      'ִ����������
    If blnUpdate Then
        blnAllow = UpdateZLHIS(strComputerName)
    End If
    CheckAllowByTerminal = blnAllow
    Exit Function
Errhand:
    MsgBox "���������ִ���" & Err.Description & "��������ϵϵͳ����Ա���н����", vbInformation, gstrSysName
End Function

Public Function StartHisCrust(ByVal str�������� As String, ByVal strJobName As String, Optional ByVal lngWait As Long, Optional ByVal strPass As String) As Boolean
'���ܣ������Զ��������
'������str��������=����ֱ�Ӵ�����ļ�·����Ҳ���Դ��ļ���
'      strJobName=�������ƣ����ߵ��ó�����
'      lngWait=��ʽ����ʱ���ȴ���N���Ӻ����ʽ����
'���أ��Ƿ�ɹ�
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
    
    If objFile.GetDriveName(str��������) = "" Then
        strUPFile = gstrSetupPath & "\" & str��������
    Else
        strUPFile = str��������
        strFileName = objFile.GetFileName(str��������)
    End If
    If Not objFile.FileExists(strUPFile) Then
        MsgBox "û���ҵ��ͻ����Զ���������" & strFileName & "������ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
        Exit Function
    End If
    If IsDesinMode Then
        '��װ�����У��Լ�����������У��λ
        strCommand = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=" & gobjRelogin.ServerName & _
                                   """;Persist Security Info=True;User ID=" & gobjRelogin.InputUser & ";Password=HIS;Data Provider=MSDASQL"
    Else
        '��װ�����У��Լ�����������У��λ
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
        MsgBox "�޷����������������̣���ʹ�ò���ϵͳ����Ա�����������", vbInformation, gstrSysName
    End If
End Function

Private Function AnalyseConfigure() As String
    '��д��:���� 2003-03-09
    '����:���������������ã�IP����������CPU���ڴ桢Ӳ�̡�����ϵͳ��
    Dim strCPU As String           'CPU
    Dim strMemory As String        '�ڴ�
    Dim strOS As String            '����ϵͳ
    Dim strComputerName As String  '�������
    Dim strHD As String            'Ӳ��
    Dim strIp As String            'IP��ַ
    Dim verinfo As OSVERSIONINFOEX
    Dim sysinfo As SYSTEM_INFO
    Dim memsts As MEMORYSTATUS
    Dim memstsex As MEMORYSTATUSEX
    Dim lngmemory As Long
    Dim curMemory As Currency
    
    strIp = LoacalIP
    '��ȡ�������
    strComputerName = ComputerName
    '��ȡӲ����Ϣ
    strHD = AnalyseHardDisk
    ' ��ò���ϵͳ��Ϣ
    strOS = GetVersionInfo
    ' ���CPU����
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
    ' ���ʣ���ڴ�
    '���ж�ϵͳ�Ƿ�Ϊwin2000������
    '�����Windows2000�����°汾������GlobalMemoryStatusȡ
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
    '��д��:���� 2003-03-09
    '����:��ȡӲ��������
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
    '�����Windows2000�����°汾��������API��ȡһ��
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
    ' ���CPU����
    GetSystemInfo sysinfo
    With myOS
        Select Case .dwMajorVersion
            Case 3
                strOS = "Windows NT 3.1"
            Case 4
                Select Case .dwMinorVersion
                    Case 0
                        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
                            strOS = "Windows NT 4.0" '1996��7�·���
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
                        strOS = "Windows 2000" '1999��12�·���
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
                        strOS = "Windows XP" '2001��8�·���
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
                            strOS = "Windows Server 2003" '2003��3�·���
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
    '����Ƿ����ظ���¼
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strProgram As String
    On Error GoTo Errhand
    
    strProgram = GetCallEXE
    strSQL = "Select A.UserName, A.Program, B.IP" & vbNewLine & _
            "From gv$Session A, zlClients B" & vbNewLine & _
            "Where A.Terminal = B.����վ" & vbNewLine & _
            "      And A.Terminal = (Select Terminal From v$Session Where AudsID = Userenv('SessionID') and RowNum =1)" & vbNewLine & _
            "      And A.Program =[1] And A.AudsID <> Userenv('SessionID')" & vbNewLine & _
            "      And B.IP <> [2]"

    Set rsTemp = OpenSQLRecord(strSQL, "����ظ�����վ", strProgram, strIpAddress)
    If rsTemp.RecordCount = 0 Then '���Ե�¼
        CheckRepeatLogin = False
        Exit Function
    Else
        MsgBox "�������д�����ͬ���Ƶļ������¼," & vbCrLf & "�Է�IP��:[" & NVL(rsTemp!IP) & "]", vbInformation, gstrSysName
        CheckRepeatLogin = True
        Exit Function
    End If
    Exit Function
Errhand:
    MsgBox "���ͬ�����������" & Err.Description & ",����ϵ������Ա���н����", vbInformation, gstrSysName
End Function

Private Function GetCallEXE() As String
'���ܣ���ȡ���õ�ǰDLL��EXE����
    Dim strPName As String, strFileName As String

    strPName = String(256, Chr(0))
    Call GetModuleFileName(0, strPName, 256)
    strFileName = Left(strPName, InStr(strPName, Chr(0)) - 1)
    strFileName = UCase(Mid(strFileName, InStrRev(strFileName, "\") + 1))
    GetCallEXE = strFileName
End Function

Private Function GetLISStation() As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'����   �õ������°�LIS��վ��
'����   �õ�վ���վ������  ��Ϊû��վ��
'        �е���֯��ʽΪ ,1,2;,վ��1,վ��2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strվ��  As String, strվ������ As String
    
    On Error GoTo Errhand
    '�ж��Ƿ������װ
    strSQL = "select 1 ���� from zlsystems where ��� = 2500 and ����� is null"
    Set rsTmp = OpenSQLRecord(strSQL, "����Ƿ������װ�°�LIS")
    If rsTmp.EOF Then Exit Function
    '�����Ƿ���Ĭ�ϵ�վ��
    strSQL = "Select Distinct A.վ��, B.����" & vbNewLine & _
            "From (Select Distinct A.վ��" & vbNewLine & _
            "       From ����������¼ A, ����������Ա B, ��Ա�� C,�ϻ���Ա�� d" & vbNewLine & _
            "       Where A.Id = B.����id And A.վ�� Is Not Null And B.��Աid = C.Id and c.id = d.��ԱID And d.�û��� = [1]) A, Zlnodelist B" & vbNewLine & _
            "Where A.վ�� = B.���" & vbNewLine & _
            "Order By A.վ��"
    Set rsTmp = OpenSQLRecord(strSQL, "վ���ѯ", gobjRelogin.DBUser)
    Do While Not rsTmp.EOF
        strվ�� = strվ�� & "," & rsTmp!վ��
        strվ������ = strվ������ & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    If strվ�� <> "" Then
        GetLISStation = strվ�� & ";" & strվ������
    End If
    Exit Function
Errhand:
    MsgBox "��ȡLIS����վ����" & Err.Description & ",����ϵ������Ա���н����", vbInformation, gstrSysName
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
'���ܣ�����ZLHIS��������
'      blnBrwCall=�Ƿ񵼺�̨����,����̨��������ʱ���Ԥ����ʱ��
    Dim strUpdateExe As String, strUpdateExePath As String
    Dim objFSO As New FileSystemObject
    Dim objConn As clsConnect, datCur           As Date
    Dim rsTemp As ADODB.Recordset, strSQL       As String
    Dim strJobName As String, blnDownload       As Boolean
    Dim strTmpPath As String, lngWait           As Long
    Dim strTmpGet  As String, blnMustNowUpdate  As Boolean
    
    strUpdateExe = "zlHisCrust.exe"
    gstrSetupPath = App.Path
    Call SaveSetting("ZLSOFT", "����ȫ��", "��������", UCase(strUpdateExe)) '����ZLRegister�������ж�
    '����Ȩ�ĳ����Զ�����
    strTmpGet = IIf(gobjRelogin.IsTransPwd, TranPasswd(gobjRelogin.InputPwd), gobjRelogin.InputPwd)
    If strTmpGet Like "δ��Ȩ�ĳ���:*" Then
        UpdateZLHIS = True
        Exit Function
    End If
    'û�з��������û��ļ��嵥��������
    If Not IsHaveClientUpgradeSet(blnForceUpdate) Then '�ͻ����޸�ʱ��������Ϣ��ʾ��
        UpdateZLHIS = True
        Exit Function
    End If
    'û���������ռ����������Զ��˳�����
    If Not CheckJobs(strComputerName, strJobName, blnBrwCall, blnForceUpdate, blnMustNowUpdate) Then
        If blnForceUpdate Then
            MsgBox "��ǰֻ�ܽ���Ԥ�������޷����пͻ����޸���", vbInformation, gstrSysName
        Else
            UpdateZLHIS = True
        End If
        Exit Function
    End If
    
    If strJobName = "OfficialUpgrade" And blnBrwCall Then
        If blnMustNowUpdate Then
            MsgBox "��⵽ϵͳ��Ҫ������Ҫ�ĸ��£�1���Ӻ������������뼰ʱ����������д�����ݡ�", vbInformation, gstrSysName
            lngWait = 1 '���������ȴ�ʱ��
        Else
            If MsgBox("��⵽ϵͳ��Ҫ�������Ƿ���������?" & vbNewLine & "ѡ���������µ�¼����������", vbInformation + vbYesNo, gstrSysName) = vbNo Then
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
    '�������򲻴��ڣ���׼������
    If Not objFSO.FileExists(strUpdateExePath) Then
        '��׼����ʱ����Ŀ¼
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
            MsgBox "�޷����ӿͻ�������������""" & objConn.ServerPath & """,����ϵ����Ա��", vbExclamation, gstrSysName
            Exit Function
        End If
        blnDownload = objConn.DownloadFile("ZLHISCRUST.EXE", strTmpPath)
        If blnDownload Then
            On Error Resume Next
            '���������ļ�
            If objFSO.FileExists(strUpdateExePath) Then
                If FileSystem.GetAttr(strUpdateExePath) <> vbNormal Then
                     Call FileSystem.SetAttr(strUpdateExePath, vbNormal)
                End If
                Call objFSO.DeleteFile(strUpdateExePath)
            End If
            If Err.Number <> 0 Then Err.Clear
            '�ȸ��Ƶ�APPSOFT�£����ʧ�ܣ����Ƶ�APPLY��
            objFSO.CopyFile strTmpPath, strUpdateExePath, True
            If Err.Number <> 0 Then
                Err.Clear
                If IsDesinMode Then
                    strUpdateExePath = "C:\APPSOFT\APPLY\zlHisCrust.exe"
                Else
                    strUpdateExePath = gstrSetupPath & "\APPLY\zlHisCrust.exe"
                End If
                '���������ļ�
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
                    '�Ƿ����°��Զ�������ǣ��ǵĻ��������ֱ�Ӵ���ʱĿ¼������
                    If UCase(GetFileDesInfo(strTmpPath, "ProductName")) = "ZLHISINSTALLUPDATE" Then
                        strUpdateExePath = strTmpPath
                    End If
                End If
            End If
        End If
        If strTmpPath <> strUpdateExePath Then
            On Error Resume Next
            '��ʱ·��
            If objFSO.FileExists(strTmpPath) Then
                If FileSystem.GetAttr(strTmpPath) <> vbNormal Then
                     Call FileSystem.SetAttr(strTmpPath, vbNormal)
                End If
                Call objFSO.DeleteFile(strTmpPath)
            End If
            Call objFSO.DeleteFolder(objFSO.GetParentFolderName(strTmpPath))
        End If
        If Not objFSO.FileExists(strUpdateExePath) Then
            MsgBox "û���ҵ��ͻ����Զ���������" & strUpdateExe & "�����޷�ͨ���������������أ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    'Ԥ���������ڵ���̨�����н���
    If StartHisCrust(strUpdateExePath, strJobName, lngWait, strTmpGet) And strJobName <> "PreUpgrade" And strJobName <> "CheckUpgrade" Then
        Exit Function
    End If
    UpdateZLHIS = True
End Function

Private Function GetDefaultFileServer() As Integer
'���ܣ���ȡĬ�Ϸ�����
'���أ���û�з��������÷���-1�����ڣ������ⷵ��һ�����������
    Dim intDefaultSever As Integer, intServerType   As Integer
    Dim blnReadOld      As Boolean
    Dim strSQL          As String, rsTmp            As ADODB.Recordset
    
    On Error Resume Next
    intDefaultSever = -1
    strSQL = "Select ��� From Zltools.Zlupgradeserver Where �Ƿ����� = 1"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ����������")
    If Err.Number <> 0 Then '���ܹ���Աʹ�õĹ�����������ͻ��˲�ƥ��
        Err.Clear
        blnReadOld = True
    ElseIf rsTmp.EOF Then
        blnReadOld = True
    End If
    On Error GoTo errH
    If Not blnReadOld Then
        intDefaultSever = Val(rsTmp!��� & "")
    Else
        strSQL = "select ���� from zlreginfo where ��Ŀ='��������'"
        Set rsTmp = OpenSQLRecord(strSQL, "���ʹ�õ���������")
        If Not rsTmp.EOF Then
            intServerType = Val(rsTmp!���� & "")
        End If
        If intServerType = 0 Then
            strSQL = "select replace(��Ŀ,'������Ŀ¼','') as ������ from zlreginfo where ��Ŀ like '������Ŀ¼%' and ���� is not null"
            Set rsTmp = OpenSQLRecord(strSQL, "����Ƿ�����õ��ļ����������")
        Else
            strSQL = "select replace(��Ŀ,'FTP������','') as ������ from zlreginfo where ��Ŀ like 'FTP������%' and ���� is not null"
            Set rsTmp = OpenSQLRecord(strSQL, "����Ƿ�����õ�FTP������")
        End If
        If Not rsTmp.EOF Then
            intDefaultSever = Val(rsTmp!������ & "")
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
        MsgBox "��ȡȱʡ����������" & Err.Description, vbInformation, gstrSysName
        Err.Clear
    End If
End Function

Private Function IsHaveClientUpgradeSet(Optional ByVal blnMsg As Boolean) As Boolean
'���ܣ��Ƿ����������ص����á�
'������blnMsg=���ΪFalse��ʱ���Ƿ���ʾ
'���أ�IsHaveClientUpgradeSet=True:���ڿ������ļ�����������ã�False-����������һ��ȱʧ
    Dim intServerID As Integer
    Dim strSQL          As String, rsTmp            As ADODB.Recordset
    
    On Error GoTo errH
    IsHaveClientUpgradeSet = True
    '���ж��Ƿ���ڿ������ļ�
    strSQL = "Select 1 �������ļ� From Zltools.Zlfilesupgrade Where Md5 Is Not Null And Rownum < 2"
    Set rsTmp = OpenSQLRecord(strSQL, "����Ƿ���ڿ������ļ�")
    If Not rsTmp.EOF Then '�������ļ���������Ҫ��һ���ж��Ƿ�����������������
        intServerID = GetDefaultFileServer
        If intServerID = -1 Then
            If blnMsg Then
                MsgBox "û�����ÿͻ��������ļ����������޷����пͻ����޸���", vbInformation, gstrSysName
            End If
            IsHaveClientUpgradeSet = False
        End If
    Else
        If blnMsg Then
            MsgBox "��δ���������ļ��嵥���޷����пͻ����޸�������ϵ����Ա��", vbInformation, gstrSysName
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
        MsgBox "��������ļ��嵥����" & Err.Description, vbInformation, gstrSysName
        Err.Clear
    End If
End Function

Private Function CheckJobs(ByVal strComputerName As String, ByRef strJobName As String, Optional ByVal blnBrwCall As Boolean, Optional ByVal blnForceUpdate As Boolean, Optional ByRef blnMustNowUpdate As Boolean) As Boolean
'����:��鲢��ȡ�������������
'      blnBrwCall=�Ƿ񵼺�̨����,����̨��������ʱ���Ԥ����ʱ��
'      blnForceUpdate=����̨����ͻ����޸�ʱ�ò���ΪTrue
'      blnMustNowUpdate=�Ƿ����ڱ�������
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim datCur As Date, blnOnlyOfficialUp As Boolean, blnOnlyPreUp As Boolean
    Dim blnPreUp As Boolean, blnOfficialUp As Boolean, blnPreComplete As Boolean, blnCollect As Boolean
    Dim strStartTime As String, strEndTime As String
    
    On Error GoTo errH
    strJobName = "": blnMustNowUpdate = False
    '���´���һ�㲻���ܳ���
    datCur = Currentdate
    '�ж������Ƿ������ȡ�Ƿ������˶�ʱ����
    strSQL = "Select Max(����) ���� From zlRegInfo Where ��Ŀ='�ͻ�����������'"
    Set rsTmp = OpenSQLRecord(strSQL, "��鶨ʱ����")
    If rsTmp!���� & "" <> "" Then
        If CDate(Format(datCur, "yyyy-MM-dd HH:mm:ss")) >= CDate(Format(NVL(rsTmp!����), "yyyy-MM-dd HH:mm:ss")) Then
            blnOnlyOfficialUp = True 'ֻ����ʽ����
        Else
            blnOnlyPreUp = True 'ֻ��Ԥ����
        End If
    Else
        blnOnlyOfficialUp = True
    End If
    On Error Resume Next
    Set rsTmp = Nothing
    '����û���Ƿ�Ԥ�����ֶ�(��ΪԤ����ʱ�����ݿ⻹û�������������Ҫ�������
    strSQL = "Select Ԥ��ʱ��,Nvl(�Ƿ�Ԥ����,0) �Ƿ�Ԥ����, Nvl(Ԥ�����, 0) Ԥ�����, Nvl(������־, 0) ������־, Nvl(�ռ���־, 0) �ռ���־,Nvl(�Ƿ���������,0) �Ƿ��������� From Zlclients Where ����վ = [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "��鵱ǰ����", strComputerName)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo errH
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            blnPreUp = rsTmp!�Ƿ�Ԥ���� = 1
            blnOfficialUp = rsTmp!������־ = 1
            blnPreComplete = rsTmp!Ԥ����� = 1
            blnCollect = rsTmp!�ռ���־ = 1
            strStartTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!Ԥ��ʱ��, "HH:00:00")
            strEndTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!Ԥ��ʱ��, "HH:59:59")
            blnMustNowUpdate = rsTmp!�Ƿ��������� = 1
        End If
    Else
        '�����·�ʽ��ȡ��ʧ����ʹ���Ϸ�ʽ�����Ӽ�����
        strSQL = "Select Ԥ��ʱ��,Nvl(Ԥ�����, 0) Ԥ�����, Nvl(������־, 0) ������־, Nvl(�ռ���־, 0) �ռ���־ From Zlclients Where ����վ = [1]"
        Set rsTmp = OpenSQLRecord(strSQL, "��鵱ǰ����", strComputerName)
        If Not rsTmp.EOF Then
            blnPreUp = rsTmp!������־ = 1
            blnOfficialUp = rsTmp!������־ = 1
            blnPreComplete = rsTmp!Ԥ����� = 1
            blnCollect = rsTmp!�ռ���־ = 1
            strStartTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!Ԥ��ʱ��, "HH:00:00")
            strEndTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!Ԥ��ʱ��, "HH:59:59")
        End If
    End If
    '��ǰֻ�ܽ���Ԥ����
    If blnOnlyPreUp Then
        '��Ԥ��������
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
        'û��Ԥ�������񣬵������ռ�����
        ElseIf blnCollect Then
            strJobName = "CheckUpgrade"
        Else
            Exit Function
        End If
    '��ǰֻ�ܽ�����ʽ����
    ElseIf blnOnlyOfficialUp Then
        If blnForceUpdate Then
            strJobName = "Repair"
        Else
            '����ʽ��������
            If blnOfficialUp Then
                strJobName = "OfficialUpgrade"
            'û����ʽ�������񣬵������ռ�����
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
        MsgBox "���ͻ����������" & Err.Description, vbInformation, gstrSysName
        Err.Clear
    End If
End Function

Public Function DeCipher(ByVal strText As String) As String
'������ܳ���
    Const MIN_ASC = 32    '��СASCII��
    Const MAX_ASC = 126 '���ASCII�� �ַ�
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intlen As Integer, intSeedLen As Integer
    Dim intStart As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    
    If strText = "" Then Exit Function
    '������ӳ���
    intSeedLen = Asc(Mid(strText, 1, 1)) - MIN_ASC
    intlen = Len(strText)
    '���þɵ�����㷨
    If intSeedLen > 0 And intSeedLen < intlen - 3 And intSeedLen < 5 Then
        '��ȡ�������
        '������ӵ������Ϊ999
        Rnd (-1)
        Randomize (999)
        For i = 2 To 1 + intSeedLen
            intChr = Asc(Mid(strText, i, 1)) 'ȡ��ĸת���ASCII��
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
        
    '���ݽ��ܵ�����
    Rnd (-1)
    Randomize (Val(strDeText))
    strDeText = ""
    For i = intStart To intlen
        intChr = Asc(Mid(strText, i, 1)) 'ȡ��ĸת���ASCII��
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
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
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
    '���ܣ���ȡ��������
    '������
    '˵����
    '******************************************************************************************************************
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    ComputerName = strComputer
    ComputerName = Trim(Replace(ComputerName, Chr(0), ""))
End Function

Public Function LoacalIP(Optional ByRef strErr As String) As String
    '���ܣ�ͨ��API��ȡ��ʱIP
    
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
'���ܣ� ȷ����ǰģʽΪ���ģʽ
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
'���ܣ�ȥ��TAB�ַ������߿ո񣬻س������ֻ�ɵ��ո�ָ���
'������strText=�����ַ�
'         blnCrlf=�Ƿ�ȥ�����з�
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
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
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
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'               ��Ϊʹ�ð󶨱���,�Դ�"'"���ַ�����,����Ҫʹ��"''"��ʽ��
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'      cnOracle=����ʹ�ù�������ʱ����
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    Dim strSQLTmp As String, arrstr As Variant
    Dim strTmp As String, strSQLtmp1 As String
    Dim lngErrNum As Long, strErrInfo As String
    
    '������ʹ���˶�̬�ڴ������û��ʹ��/*+ XXX*/����ʾ��ʱ�Զ�����
    strSQLTmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLTmp, 7)), 1, 2) <> "/*" And Mid(strSQLTmp, 1, 6) = "SELECT" Then
        arrstr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrstr)
            strSQLtmp1 = strSQLTmp
            Do While InStr(strSQLtmp1, arrstr(i)) > 0
                '�ж�ǰ���Ƿ�����IN �����򲻼�Rule
                '���ҵ����һ��SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrstr(i)) - 1)
                strTmp = Replace(FromatSQL(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  'ȡ����3���ַ�
                
                If strTmp = "IN(" Then '����in(select��������������?��������Ƿ����û��ʹ������д����������̬�ڴ溯�?
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
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '������������"[����]����"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop
    
    If UBound(arrInput) + 1 < intMax Then
        Err.Raise 9527, strTitle, "SQL���󶨱�����ȫ��������Դ��" & strTitle
    End If

    '�滻Ϊ"?"����
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next
    
    '�����µĲ���
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '����
            '���ַ�ʽ������һЩIN�Ӿ��Union���
            '��ʾͬһ�������Ķ��ֵ,�����Ų�������������Ĳ����Ž���,��Ҫ��֤�����ֵ��������
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '�ַ�
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '�ò������������õ��ڼ���ֵ��
        End Select
    Next
'    If gblnSys = True Then
'        Set cmdData.ActiveConnection = gcnSysConn
'    Else
    Set cmdData.ActiveConnection = gcnOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
'    End If
    cmdData.CommandText = strSQL
    
'    Call gobjComLib.SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecordByArray = cmdData.Execute
    Set OpenSQLRecordByArray.ActiveConnection = Nothing
'    Call gobjComLib.SQLTest
End Function

Public Sub ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'���ܣ�ִ�й������,���Զ��Թ��̲������а󶨱�������
'������strSQL=�������,���ܴ�����,����"������(����1,����2,...)"��
'      cnOracle=����ʹ�ù�������ʱ����
'˵�������¼���������̲�����ʹ�ð󶨱���,�����ϵĵ��÷�����
'  1.���������Ǳ��ʽ,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1,100.12*0.15,...)"
'  2.�м�û�д�����ȷ�Ŀ�ѡ����,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1, , ,����3,...)"
'  3.��Ϊ�ù������Զ�����,����һ��ʹ�ð󶨱���,�Դ�"'"���ַ�����,��Ҫʹ��"''"��ʽ��
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    Dim lngErrNum As Long, strErrInfo As String
    
    If Right(Trim(strSQL), 1) = ")" Then
        'ִ�еĹ�����
        strTemp = Trim(strSQL)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        'ִ�й��̲���
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For i = 1 To Len(strTemp)
            '�Ƿ����ַ����ڣ��Լ����ʽ��������
            If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                strPar = Trim(strPar)
                With cmdData
                    If IsNumeric(strPar) Then '����
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, strPar)
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '�ַ���
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        
                        'Oracle���ӷ�����:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '˫"''"�İ󶨱�������
                        If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                        
                        '���Ӳ�������LOBʱ������ð󶨱���ת��ΪRAWʱ����2000���ַ�Ҫ��adLongVarChar
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
                        If intMax <= 2000 Then
                            intMax = IIf(intMax <= 200, 200, 2000)
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMax, strPar)
                        Else
                            If intMax < 4000 Then intMax = 4000
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adLongVarChar, adParamInput, intMax, strPar)
                        End If
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '����
                        strPar = Split(strPar, "(")(1)
                        strPar = Trim(Split(strPar, ",")(0))
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        If strPar = "" Then
                            'NULLֵ�������ִ���ɼ�����������
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(strPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , CDate(strPar))
                        End If
                    ElseIf UCase(strPar) = "SYSDATE" Then '����
                        If datCur = CDate(0) Then datCur = Currentdate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then 'NULLֵ�����ַ�����ɼ�����������
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, 200, Null)
                    ElseIf strPar = "" Then '��ѡ��������NULL������ܸı���ȱʡֵ:��˿�ѡ��������д���м�
                        GoTo NoneVarLine
                    Else '�������������ӵı��ʽ���޷�����
                        GoTo NoneVarLine
                    End If
                End With
                
                strPar = ""
            Else
                strPar = strPar & Mid(strTemp, i, 1)
            End If
        Next
        
        '����Ա���ù���ʱ��д����
        If blnStr Or intBra <> 0 Then
            Err.Raise -2147483645, , "���� Oracle ����""" & strProc & """ʱ�����Ż�������д��ƥ�䡣ԭʼ������£�" & vbCrLf & vbCrLf & strSQL
            Exit Sub
        End If
        
        '����?��
        strTemp = ""
        For i = 1 To cmdData.Parameters.Count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
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
    '˵����Ϊ�˼��������ӷ�ʽ
    '1.��������adCmdStoredProc��ʽ��8i����������
    '2.�����������ʹ��{},��ʹ����û�в���ҲҪ��()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    gcnOracle.Execute strSQL, , adCmdText
'    Call gobjComLib.SQLTest
End Sub


Public Function IP(Optional ByVal strErr As String) As String
    '******************************************************************************************************************
    '����:ͨ��oracle��ȡ�ļ������IP��ַ
    '���:strDefaultIp_Address-ȱʡIP��ַ
    '����:
    '����:����IP��ַ
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim strIp_Address As String
    Dim strSQL As String
        
    On Error GoTo Errhand
    
    strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡIP��ַ")
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
    '���ܣ���ȡ�������ϵ�ǰ����
    '������
    '���أ�����Oracle���ڸ�ʽ�����⣬����
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
'���ܣ�������븴�Ӷ�
'������cnInput=���������
'          strChcekPWD=�ȴ���������
'          strToolTip=�����ʾ����
'���أ�True-���ɹ���False-���ʧ��
    Dim strSQL As String, rsData As New ADODB.Recordset
    Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
    Dim blnPwdLen As Boolean, intPwdMin As Integer, intPwdMax As Integer
    Dim blnComplex As Boolean, strOterChrs As String
    Dim lngLen As Long, i As Integer, intChr As Integer
    
    On Error GoTo errH
    strToolTip = ""
    strSQL = "Select ������,Nvl(����ֵ,ȱʡֵ) ����ֵ From zlOptions Where ������ in (20,21,22,23)"
    rsData.Open strSQL, cnInput
    blnPwdLen = False: intPwdMin = 0: intPwdMax = 0
    blnComplex = False: strOterChrs = ""
    Do While Not rsData.EOF
        Select Case rsData!������
            Case 20 '�Ƿ�������볤��
                blnPwdLen = Val(rsData!����ֵ & "") = 1
            Case 21 '���볤������
                intPwdMin = Val(rsData!����ֵ & "")
            Case 22 '���볤������
                intPwdMax = Val(rsData!����ֵ & "")
            Case 23 '�Ƿ�������븴�Ӷ�
                blnComplex = Val(rsData!����ֵ & "") = 1
        End Select
        rsData.MoveNext
    Loop
    '����������ʾ
    If blnPwdLen Then
        If intPwdMin = intPwdMax Then
            strToolTip = "�������Ϊ" & intPwdMax & " λ�ַ���"
        Else
            strToolTip = "�������Ϊ" & intPwdMin & "��" & intPwdMax & " λ�ַ���"
        End If
     End If
     If blnComplex Then
        If strToolTip <> "" Then
            strToolTip = strToolTip & vbNewLine & "���ٰ���һ�����֡�һ����ĸ��һ�������ַ���ɡ�"
        Else
            strToolTip = "������һ�����֡�һ����ĸ��һ�������ַ���ɡ�"
        End If
     End If
    '���ȼ��
    lngLen = ActualLen(strChcekPWD)
    If lngLen <> Len(strChcekPWD) Then
        MsgBox "���������˫�ֽ��ַ������飡", vbInformation, gstrSysName
        Exit Function
    End If
    If blnPwdLen Then
        If Not (lngLen >= intPwdMin And lngLen <= intPwdMax) Then
            If intPwdMin = intPwdMax Then
                MsgBox "�������Ϊ" & intPwdMax & " λ�ַ���", vbInformation, gstrSysName
                Exit Function
            Else
                MsgBox "�������Ϊ" & intPwdMin & "��" & intPwdMax & " λ�ַ���", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    For i = 1 To Len(strChcekPWD)
        intChr = Asc(UCase(Mid(strChcekPWD, i, 1)))
        If intChr >= 32 And intChr < 127 Then
            'Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
            Select Case intChr
                Case 48 To 57 '����
                    blnHaveNum = True
                Case 65 To 90 '��ĸ
                    blnAlpha = True
                Case 32, 34, 47, 64  '�ո�,˫����,/,@
                    strOterChrs = strOterChrs & Chr(intChr)
                Case Is < 48, 58 To 64, 91 To 96, Is > 122
                    blnChar = True
            End Select
        Else
            strOterChrs = strOterChrs & Chr(intChr)
        End If
    Next
    If strOterChrs <> "" Then
        MsgBox "���벻�����������ַ���" & strOterChrs, vbInformation, gstrSysName
        Exit Function
    ElseIf Not (blnHaveNum And blnAlpha And blnChar) And blnComplex Then
        MsgBox "����������һ�����֡�һ����ĸ��һ�������ַ���ɡ�", vbInformation, gstrSysName
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
    '���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
    '       ʵ�����ݴ洢����
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function OpenImeByName(Optional strIme As String) As Boolean
'����:�����ƴ��������뷨,��ָ������ʱ�ر��������뷨��֧�ֲ������ơ�
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim blnNotCloseIme As Boolean
    
    If strIme = "���Զ�����" Then OpenImeByName = True: Exit Function
    
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
        '����windows Vistaϵͳ��Ӣ�����뷨��ImmIsIME���Գ���true�����뷨,���,��Ҫ��������.
        '���˺�:2008/09/03
        If ActivateKeyboardLayout(arrIme(0), 0) <> 0 Then OpenImeByName = True: Exit Function
    End If
End Function
