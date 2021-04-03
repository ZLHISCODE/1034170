Attribute VB_Name = "mdlClient"
Option Explicit
 
Public gcnOracle As ADODB.Connection
Public gstrServerPath As String         '������Ŀ¼
Public gstrSoftPath As String           'Ŀ��Ӧ�ò���������Ŀ¼
Public gstrWinSystemPath As String      'ϵͳĿ¼
Public gstrComputerIp As String         '������IP��ַ
Public gstrComputerName As String       '������
Public gstrAppPath  As String           '��ǰĿ¼

Public gstrVisitUser As String       '���ʵ��û���
Public gstrVisitPassWord As String   '���ʵ�����
Public gstrVisitPort As String       '���ʵĶ˿�
Public gstrConnectString As String
Public gBlnHisCrustCompare As Boolean '�Ƿ�Ƚ�HISCUST��ǳ���
Public gstrHisCommand   As String     'ZLHIS�����������ʱ����Ĳ������ص�ʱ����
Public gstrWinPath As String           'windowsĿ¼
Public gstrAppEXE   As String           '���ñ���ǳ�����ļ�
Public gblnPreUpgrade As Boolean        '�Ƿ�ΪԤ����
Public gblnOfficialUpgrade As Boolean   '�Ƿ�Ϊ��ʽ����
Public gblnԤ����� As Boolean          'Ԥ�����Ƿ����

Public gstr�ռ����� As String           '��:log;doc��
Public gbln�ռ� As Boolean
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

'��ʾ��ǰ���еĴ����API����
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GWL_STYLE = (-16)
Public Const WS_VISIBLE = &H10000000
Public Const WS_BORDER = &H800000
Private Const INTERNET_FLAG_PASSIVE = &H8000000     '���ñ�������

'�رյ�ǰ���еĴ����API����
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


'ȡIP��API
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
  ComboIndex            As Long  '����
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
Public gintUpType  As Integer '������ʽ 0���� 1FTP
Public gintGatherTYpe As Integer '�ռ���ʽ 0���� 1FTP
Public gblnOk      As Boolean
Public gstrTempPath As String  '��ʱ���Ŀ¼
Public gstrPerTempPath As String 'Ԥ������ʱ·��

Private Sub Main()
    Dim arrCommand As Variant
'Command =
'Ԥ������   ConnectionString||��ǵ���(1-��ǵ��õ�,0��������õ�)||PreUpgrade||�����в���
'��ʱ������ ConnectionString||��ǵ���(1-��ǵ��õ�,0��������õ�)||OfficialUpgrade���·����ZLHIS+ִ���ļ�||�����в���||USER=ZLHIS PASS=HIS(�������������)

'ConnectionString
'10.35.10��ǰ��Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=30_TXYY_135"";Persist Security Info=True;User ID=zlhis;Password=HIS;Data Provider=MSDASQL
'10.35.10���Ժ�Provider=MSDataShape.1;Persist Security Info=False;User ID=ZLHIS;Data Provider=MSDASQL; ���Ӵ��в�������ͷ���������

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

    
    '�Ƿ�Ԥ����
    gblnOfficialUpgrade = False
    gblnPreUpgrade = (gstrAppEXE = "PreUpgrade")
    
    '��ʽ����
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
'����:ȷ���Ƿ�������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    IsUpgrade = False
    gbln�ռ� = False
    
    If gblnPreUpgrade Then
        strSQL = "Select 1 From zlClients Where  Ԥ�����=1 and ����վ=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "���Ԥ��", gstrComputerName)
        IsUpgrade = rsTmp.RecordCount = 0
        Exit Function
    Else
        strSQL = "Select 1 From zlClients Where  ������־=1 and ����վ=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "����Ƿ�����", gstrComputerName)
        If rsTmp.RecordCount = 0 Then
            strSQL = "Select 1 From zlClients Where  �ռ���־=1 and ����վ=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, "����Ƿ��ռ�", gstrComputerName)
            If rsTmp.RecordCount <> 0 Then
                gbln�ռ� = True
                IsUpgrade = True
            End If
        Else
            IsUpgrade = True
        End If
    End If
    
End Function

Public Function GetVersion(ByVal lngVer As Variant) As String
    '����:������ת���ɰ汾��
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
    '����:�����Ƚ�(�Ƚϰ汾��,�޸�ʱ��)
    '�����:
    '   strSourceFile:Դ�ļ�
    '   strTargetFile:Ŀ���ļ�
    '������:
    '   strSourceVer:Դ�汾��
    '   strSourceDate:Դ�ļ�������޸�����
    '   strTargetVer:Ŀ��汾��
    '   strTargetDate:Ŀ���ļ�������޸�����
    '����:��Դ�ļ���Ŀ���ļ���,�򷵻�True,���򷵻�False

    Dim objFile As New FileSystemObject
    Dim strSource As String
    Dim strTarget As String
    
    CompareFile = False

    On Error Resume Next
    
    '�Ƚ��ļ��汾��
    strSource = strSourceVer
    strTarget = GetCommpentVersion(strTargetFile)
    
    strSourceVer = strSource
    strTargetVer = strTarget
    If RtnVerNum(strTarget) < RtnVerNum(strSource) Then
        CompareFile = True
        
    End If
    
    
    '�Ƚ��ļ�������޸�ʱ��
    
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
    , Optional ByRef strErrMsg As String = "����������", Optional notReg As Boolean = False, Optional blnSysFile As Boolean = False, Optional notCopy As Boolean = False) As Boolean

    '����:��Դ�Ŀ���Ŀ���ļ�
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-01-20 13:46:50

    Dim objFile As New FileSystemObject, cllProcess As New Collection   '���̼�array(����,Exe�ļ���,ģ�����)
    Dim strFile As String, intType As Integer '0-Exe��ִ���ļ�,1-DLL;OCX��Ҫע����ļ�,2-�����ļ�,��CHM,log���ļ�
    Dim i As Long, strMsgBox As String, iTask As Integer
    Dim pHandle As Long, ret As Integer
    Dim regMsg
    Dim strResult As String
    Dim lngResult As Long
    
    strFile = UCase(strSourceFile)
    If strFile Like "*.EXE" Then
        '�ж��Ƿ�ΪACTIVEX EXE���,�Ƿ���Ҫע��
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
    
    '������ļ�����=5,��ǿ������,���ұ����ļ�����,ֱ���ж�Ϊ�ɹ�����!
    If blnSysFile = True And notCopy = False And objFile.FileExists(strTargetFile) = True Then
        GoTo GoSuress
    End If
    
'    'Scrrun.dll �����ļ�,���⴦��
'    If InStrRev(UCase(strSourceFile), UCase("\scrrun.dll")) > 0 Then
'        If objFile.FileExists(strTargetFile) Then
'            GoTo GoSuress
'        End If
'    End If
    
    '��ע��
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
    
    
    '����ֻ���ļ�,���������Ϊ��ͨ����
    If objFile.FileExists(strTargetFile) Then
        '�ļ�����,�������
        If FileSystem.GetAttr(strTargetFile) <> vbNormal Then
                FileSystem.SetAttr strTargetFile, vbNormal
        End If
    End If
    
    '��鱾��Ŀ¼�Ƿ����
    Call CreatePath(strTargetFile)
    

    '�����ļ�
    objFile.CopyFile strSourceFile, strTargetFile, True
    If Err <> 0 Then
        '����Ƿ�Ϊϵͳ�ļ�
      
       strErrMsg = Err.Number & "-" & Err.Description

       '�ܾ���Ȩ��
       If CheckSysFile(strTargetFile) Then
           '�ܾ�Ȩ���ȸ���
           
           If notCopy Then 'ǿ�Ƹ���
                Err.Clear
                
                On Error Resume Next
                Call Kill(strTargetFile & "_old")
                Name strTargetFile As strTargetFile & "_old"
                Call Kill(strTargetFile & "_old")
                
                '���¿����ļ�
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
       
       '��������,�϶������ļ���ֻ���򱻶�ռ�򿪻���ִ��
       If Err.Number <> 70 And Err.Number <> 70 - 2146828288 Then
            If MsgBox("ע�⣺" & vbCrLf & _
                       "     �ļ���" & strTargetFile & "������������ ,ԭ�����£�" & vbCrLf & Err.Number & "-" & Err.Description & vbCrLf & _
                       "�����ԡ���ʾ�ֹ��Ѿ������ش�������ִ��������" & vbCrLf & _
                       "��ȡ������ʾȡ������������", vbQuestion + vbRetryCancel + vbDefaultButton1, "�Զ�����") = vbRetry Then
               '����ִ��һ�ο���
               GoTo GoReExectue:
            Else
               Exit Function
            End If
       End If

        Select Case intType
        Case 0  ''0-Exe��ִ���ļ�
             Call zlGetFileProcess(strFile, cllProcess)
             strMsgBox = ""
             If gbln�ռ� Then
                 regMsg = MsgBox("ע�⣺" & vbCrLf & _
                        "     �ļ���" & strTargetFile & "������û��дȨ�ޣ������������������ռ� ��" & vbCrLf & _
                        "����ֹ����ʾȡ�����ļ��ռ���" & vbCrLf & _
                        "�����ԡ���ʾ�ֹ��Ѿ�����������Ŀ¼Ȩ�ޣ�����ִ���ռ���" & vbCrLf & _
                        "�����ԡ���ʾ���β������ռ���", vbQuestion + vbAbortRetryIgnore, "�Զ��ռ�")
             Else
                 regMsg = MsgBox("ע�⣺" & vbCrLf & _
                        "     �ļ���" & strTargetFile & "������ִ�У�����������" & vbCrLf & _
                        "����ֹ����ʾȡ��������������" & vbCrLf & _
                        "�����ԡ���ʾ��ֹ�����еĳ�������ִ��������" & vbCrLf & _
                        "�����ԡ���ʾ���β�����������", vbQuestion + vbAbortRetryIgnore, "�Զ�����")    'vbAbortRetryIgnore
             End If
             
             If regMsg = 3 Then
                Exit Function
             ElseIf regMsg = 4 Then
                '�Ƚ�����صĽ���
                For i = 0 To cllProcess.Count
                    Call TerminatePID(cllProcess(i)(0))
                Next
                '����ִ��һ�ο���
                GoTo GoReExectue:
             Else
                strErrMsg = "���Ա���������"
                GoTo GoSuress
             End If
        Case 1  '1-DLL;OCX��Ҫע����ļ�
            strMsgBox = ""
            Call zlGetFileProcess(strFile, cllProcess)
            For i = 1 To cllProcess.Count
                If UCase(cllProcess(i)(1)) = UCase("ZLHISCRUST.EXE") Then
                    strErrMsg = "�������ռ"
                    GoTo GoSuress
                End If
                If i > 2 Then
                    strMsgBox = strMsgBox & Space(5) & cllProcess(i)(0) & "��" & cllProcess(i)(1) & vbCrLf & Space(5) & "...."
                    Exit For
                Else
                    strMsgBox = strMsgBox & Space(5) & cllProcess(i)(0) & "��" & cllProcess(i)(1) & vbCrLf
                End If
            Next
            '3��ֹ
            '4����
            '5����
            If gbln�ռ� Then
                regMsg = MsgBox("ע�⣺" & vbCrLf & _
                        "     �ļ���" & strTargetFile & "������û��дȨ�ޣ������������������ռ� ��" & vbCrLf & _
                        "����ֹ����ʾȡ�����ļ��ռ���" & vbCrLf & _
                        "�����ԡ���ʾ�ֹ��Ѿ�����������Ŀ¼Ȩ�ޣ�����ִ���ռ���" & vbCrLf & _
                        "�����ԡ���ʾ���β������ռ���", vbQuestion + vbAbortRetryIgnore, "�Զ��ռ�")
            Else
                regMsg = MsgBox("ע�⣺" & vbCrLf & _
                        "     �ļ���" & strTargetFile & "���������³������ã��������� ��" & vbCrLf & _
                        strMsgBox & vbCrLf & _
                        "����ֹ����ʾȡ��������������" & vbCrLf & _
                        "�����ԡ���ʾ��ֹ�����еĳ�������ִ��������" & vbCrLf & _
                        "�����ԡ���ʾ���β�����������", vbQuestion + vbAbortRetryIgnore, "�Զ�����")    'vbAbortRetryIgnore
            End If
            
            If regMsg = 3 Then
                Exit Function
            ElseIf regMsg = 4 Then
               '�Ƚ�����صĽ���
               For i = 0 To cllProcess.Count
                   Call TerminatePID(cllProcess(i)(0))
               Next
               '����ִ��һ�ο���
               GoTo GoReExectue:
            Else
               strErrMsg = "���Ա���������"
               GoTo GoSuress
            End If
        Case Else   ',2-�����ļ�,��CHM,log���ļ�
            If gbln�ռ� Then
                regMsg = MsgBox("ע�⣺" & vbCrLf & _
                        "     �ļ���" & strTargetFile & "������û��дȨ�ޣ������������������ռ� ��" & vbCrLf & _
                        "����ֹ����ʾȡ�����ļ��ռ���" & vbCrLf & _
                        "�����ԡ���ʾ�ֹ��Ѿ�����������Ŀ¼Ȩ�ޣ�����ִ���ռ���" & vbCrLf & _
                        "�����ԡ���ʾ���β������ռ���", vbQuestion + vbAbortRetryIgnore, "�Զ��ռ�")
            Else
                regMsg = MsgBox("ע�⣺" & vbCrLf & _
                        "     �ļ���" & strTargetFile & "�����������ļ���վ�򿪣��������� ��" & vbCrLf & _
                        "����ֹ����ʾȡ��������������" & vbCrLf & _
                        "�����ԡ���ʾ�ֹ��Ѿ������վ���еĳ�������ִ��������" & vbCrLf & _
                        "�����ԡ���ʾ���β�����������", vbQuestion + vbAbortRetryIgnore, "�Զ�����")
            End If
            
            If regMsg = 3 Then
               Exit Function
            ElseIf regMsg = 4 Then
               '����ִ��һ�ο���
               GoTo GoReExectue:
            Else
               strErrMsg = "���Ա���������"
               GoTo GoSuress
            End If
        End Select
    End If
Redo:
    If intType = 1 And notReg Then
        '1-DLL;OCX��Ҫע����ļ�,
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
    
    strErrMsg = "��������!"
GoSuress:
    GetCopyAndReg = True
End Function

Public Function GetWinPath() As String
    '--����:��ȡϵͳĿ¼
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
    Dim str�������� As String
    
    On Error GoTo errH
    
    '��ʼ������
    gstrWinSystemPath = GetWinSystemPath()
    gstrWinPath = GetWinPath()
    gstrSoftPath = GetSoftPath()
    
    If IsSourceCode Then
        gstrAppPath = "C:\APPSOFT"
    Else
        gstrAppPath = Replace(UCase(App.Path), UCase("\Apply"), "", 1)
    End If
    
    gstrComputerIp = AnalyseIP
    '��ȡ����վ��
    gstrComputerName = GetMyCompterName
    
    If gblnPreUpgrade Then
        strSQL = "Select ����������,FTP������ From zlClients Where ����վ=[1]"
    Else
        strSQL = "Select ����������,FTP������ From zlClients Where  ������־=1 and ����վ=[1]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "���Ԥ��", gstrComputerName)
    
    With rsTmp
        If .RecordCount = 0 Then
            gbln�ռ� = Not gblnPreUpgrade
        Else
            gbln�ռ� = False
            If gintUpType = 0 Then
                If IsNull(rsTmp!����������) Then
                    str�������� = ""
                Else
                    str�������� = rsTmp!����������
                End If
            Else
                If IsNull(rsTmp!FTP������) Then
                    str�������� = ""
                Else
                    str�������� = rsTmp!FTP������
                End If
            End If
        End If
    End With
    
    
    If gbln�ռ� Then
        If gintGatherTYpe = 0 Then
            str�������� = "S"
            strSQL = "Select ��Ŀ,���� From zlregInfo where ��Ŀ in('�ռ�Ŀ¼S','�����û�S','��������S','�ռ�����')"
        Else
            str�������� = "F"
            strSQL = "Select ��Ŀ,���� From zlregInfo where ��Ŀ in('�ռ�Ŀ¼F','�����û�F','��������F','���ʶ˿�F','�ռ�����')"
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "��ʼ��")
        With rsTmp
            Do While Not .EOF
                If !��Ŀ = "�ռ�Ŀ¼" & str�������� Then
                    gstrServerPath = IIf(IsNull(!����), "", !����)
                End If
                If !��Ŀ = "�����û�" & str�������� Then
                    gstrVisitUser = IIf(IsNull(!����), "", !����)
                End If
                If !��Ŀ = "��������" & str�������� Then
                    gstrVisitPassWord = IIf(IsNull(!����), "", !����)
                End If
                If !��Ŀ = "���ʶ˿�" & str�������� Then
                    gstrVisitPort = IIf(IsNull(!����), "", !����)
                End If
                If !��Ŀ = "�ռ�����" Then
                    gstr�ռ����� = IIf(IsNull(!����), "", !����)
                End If
                .MoveNext
            Loop
        End With
    Else
        '��������
        If gintUpType = 0 Then
            'Share����ʽ
            If str�������� = "" Then
                strSQL = "Select ��Ŀ,���� From zlregInfo where ��Ŀ in('������Ŀ¼','�����û�','��������')"
            Else
                strSQL = "Select ��Ŀ,���� From zlregInfo where ��Ŀ in('������Ŀ¼" & str�������� & "','�����û�" & str�������� & "','��������" & str�������� & "')"
            End If
            Set rsTmp = OpenSQLRecord(strSQL, "��ʼ��")
            With rsTmp
                Do While Not .EOF
                    If !��Ŀ = "������Ŀ¼" & str�������� Or !��Ŀ = "�ռ�Ŀ¼" Then
                        gstrServerPath = IIf(IsNull(!����), "", !����)
                    End If
                    If !��Ŀ = "�����û�" & str�������� Then
                        gstrVisitUser = IIf(IsNull(!����), "", !����)
                    End If
                    If !��Ŀ = "��������" & str�������� Then
                        gstrVisitPassWord = IIf(IsNull(!����), "", !����)
                    End If
                    If !��Ŀ = "�ռ�����" Then
                        gstr�ռ����� = IIf(IsNull(!����), "", !����)
                    End If
                    .MoveNext
                Loop
            End With
            
            '20101206-zq ���⴦�������0 �� NULL ֵΪһ��,�������������������������
            If gstrServerPath = "" And gbln�ռ� = False Then
                If str�������� = 0 Then
                     strSQL = "Select ��Ŀ,���� From zlregInfo where ��Ŀ in('������Ŀ¼','�����û�','��������')"
                     Set rsTmp = OpenSQLRecord(strSQL, "��������")

                     With rsTmp
                        Do While Not .EOF
                            If !��Ŀ = "������Ŀ¼" Or !��Ŀ = "�ռ�Ŀ¼" Then
                                gstrServerPath = IIf(IsNull(!����), "", !����)
                            End If
                            If !��Ŀ = "�����û�" Then
                                gstrVisitUser = IIf(IsNull(!����), "", !����)
                            End If
                            If !��Ŀ = "��������" Then
                                gstrVisitPassWord = IIf(IsNull(!����), "", !����)
                            End If
                            If !��Ŀ = "�ռ�����" Then
                                gstr�ռ����� = IIf(IsNull(!����), "", !����)
                            End If
                            .MoveNext
                        Loop
                        
                     End With
                End If
            End If
        Else
            'FTP����ʽ
            If str�������� = "" Then str�������� = "0"
            strSQL = "Select ��Ŀ,���� From zlregInfo where ��Ŀ in('FTP������" & str�������� & "','FTP�û�" & str�������� & "','FTP����" & str�������� & "','FTP�˿�" & str�������� & "')"
            Set rsTmp = OpenSQLRecord(strSQL, "���FTP")
            
            With rsTmp
                Do While Not .EOF
                    If !��Ŀ = "FTP������" & str�������� Or !��Ŀ = "�ռ�Ŀ¼" Then
                        gstrServerPath = IIf(IsNull(!����), "", !����)
                    End If
                    If !��Ŀ = "FTP�û�" & str�������� Then
                        gstrVisitUser = IIf(IsNull(!����), "", !����)
                    End If
                    If !��Ŀ = "FTP����" & str�������� Then
                        gstrVisitPassWord = IIf(IsNull(!����), "", !����)
                    End If
                    If !��Ŀ = "FTP�˿�" & str�������� Then
                        gstrVisitPort = IIf(IsNull(!����), "", !����)
                    End If
                    If !��Ŀ = "�ռ�����" Then
                        gstr�ռ����� = IIf(IsNull(!����), "", !����)
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
'--��  ��:�����û���������
'--�����:
'--������:
'--��  ��:

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
            'ȡ�û���
            strUserName = Mid(strArr(i), j + 8)
        End If
        j = InStr(1, UCase(strArr(i)), UCase("Password"))
        If j <> 0 Then
            'ȡ�û���
            strPass = Mid(strArr(i), j + 9)
        End If
    Next
    strUser = strUserName
    strPassWord = strPass

End Sub

Public Function IsNetServer() As Boolean
    '--����:���������Ƿ�����������
    Dim NetR As NETRESOURCE
    Dim objFile As New FileSystemObject
      
    '���˺�:���ܴ���windows��Դ�������Ѿ��з��ʵ���
    '
    If objFile.FolderExists(gstrServerPath) Then
            IsNetServer = True: Exit Function
    End If
    
    If objFile.FolderExists(gstrServerPath) Then '���ڴ��ļ���,�϶�û��Ȩ�޷���,��Ҫɾ������
            Call zlNetCancelConnected 'Ŀǰȫ��ɱ��,ԭ���ǲ�֪���ļ���������:��:IP�ͻ���������
    End If
    
    
    With NetR
        .dwScope = RESOURCE_GLOBALNET
        .dwType = RESOURCETYPE_DISK
        .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
        .dwUsage = RESOURCEUSAGE_CONNECTABLE
        .lpLocalName = "" 'ӳ���������
        .lpRemoteName = gstrServerPath  '������·��
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
'����:����Ƿ�����������FTP������
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
    '�Ͽ�����������
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
    
    '����:��ȡӦ�ò���������Ŀ¼
    
    Dim strFile As String
    Dim strPath As String
    Dim objFile As New FileSystemObject
    Dim blnRutn As Boolean
    Dim DirPath  As String
    Dim ll As FileListBox
    
    On Error GoTo ErrHand:
    '---ȷ����ǰϵͳĿ¼
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
    '--����:����ָ�����ļ��Ƿ����
    '--����: ������ڴ��ļ�ΪTrue,����ΪFlase
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
    '����:������������IP��ַ
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

    '--����:�������ְ汾

    Dim strArr
    
    If strVer <> "" Then
        strArr = Split(strVer, ".")
        RtnVerNum = strArr(0) * 10 ^ 8 + strArr(1) * 10 ^ 4 + strArr(2)
    Else
        RtnVerNum = 0
    End If
End Function

Public Function isHisCurstUpGrade() As Boolean

    '--����:ȷ���Ƿ��������,��������ǿ���applyĿ¼��

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
    strSQL = "Select �޸�����,�汾��,MD5 From zlfilesupgrade where upper(�ļ���)='ZLHISCRUST.EXE'"
    Set rsTmp = OpenSQLRecord(strSQL, "���������")
    
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
'            strSourceFile = "D:\zlHisCrust.exe" ����
            strSourceMD5 = HashFile(strSourceFile, 2 ^ 27)
        Else
            strSourceVer = "0.0.0"
            strSourceDate = ""
            strSourceMD5 = ""
        End If
        
        strTargetMD5 = NVL(!MD5, "")
         
         '�Ƚ�MD5
         If UCase(strSourceMD5) <> UCase(strTargetMD5) Then
            strTargetVer = GetVersion(IIf(IsNull(!�汾��), 0, !�汾��))
            strTargetDate = Format(!�޸�����, "yyyy-MM-DD HH:mm:ss")
            
            '�ж�Ŀ¼�Ƿ����,�������Զ�����.20101206--ZQ
            If mobjFile.FolderExists(gstrAppPath & "\Apply\") = False Then
               Call mobjFile.CreateFolder(gstrAppPath & "\Apply\")
            End If
            
            If gintUpType = 0 Then
                If GetCopyAndReg(strTargetFile, gstrAppPath & "\Apply\zlHisCrust.exe", strErrMsg) Then
                    isHisCurstUpGrade = True
                    WriteTxtLog strSourceFile & "(�汾:" & strSourceVer & "   �޸�����:" & strSourceDate & ")    ====>    " & _
                                    strTargetFile & "(�汾:" & strTargetVer & "   �޸�����:" & strTargetDate & ")        ������Ϣ:" & strErrMsg
                    Exit Function
                End If
            Else
                If FtpDownFile(strTargetFile, gstrAppPath & "\Apply\zlHisCrust.exe") Then
                    isHisCurstUpGrade = True
                    WriteTxtLog strSourceFile & "(�汾:" & strSourceVer & "   �޸�����:" & strSourceDate & ")    ====>    " & _
                                    strTargetFile & "(�汾:" & strTargetVer & "   �޸�����:" & strTargetDate & ")        ������Ϣ:" & strErrMsg
                    Exit Function
                End If
            End If
        End If
        
    End With
    Exit Function
ErrHand:
End Function

Public Function isRegComUpGrade(ByVal strRegName As String) As Boolean

    '--����:ȷ���Ƿ��������,��������ǿ���applyĿ¼��

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
    strSQL = "Select �޸�����,�汾��,MD5 From zlfilesupgrade where upper(�ļ���)='" & strRegName & "'"
    Set rsTmp = OpenSQLRecord(strSQL, "���������")
    
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
'            strSourceFile = "D:\zlHisCrust.exe" ����
            strSourceMD5 = HashFile(strSourceFile, 2 ^ 27)
        Else
            strSourceVer = "0.0.0"
            strSourceDate = ""
            strSourceMD5 = ""
        End If
        
        strTargetMD5 = NVL(!MD5, "")
         
         '�Ƚ�MD5
         If UCase(strSourceMD5) <> UCase(strTargetMD5) Then
            strTargetVer = GetVersion(IIf(IsNull(!�汾��), 0, !�汾��))
            strTargetDate = Format(!�޸�����, "yyyy-MM-DD HH:mm:ss")
            
            '�ж�Ŀ¼�Ƿ����,�������Զ�����.20101206--ZQ
            If mobjFile.FolderExists(gstrAppPath & "\PUBLIC\") = False Then
               Call mobjFile.CreateFolder(gstrAppPath & "\PUBLIC\")
            End If
            
            If gintUpType = 0 Then
                If GetCopyAndReg(strTargetFile, gstrWinSystemPath & "\" & strRegName, strErrMsg) Then
                    isRegComUpGrade = True
                    WriteTxtLog strSourceFile & "(�汾:" & strSourceVer & "   �޸�����:" & strSourceDate & ")    ====>    " & _
                                    strTargetFile & "(�汾:" & strTargetVer & "   �޸�����:" & strTargetDate & ")        ������Ϣ:" & strErrMsg
                    Exit Function
                End If
            Else
                If FtpDownFile(strTargetFile, gstrWinSystemPath & "\" & strRegName) Then
                    isRegComUpGrade = True
                    WriteTxtLog strSourceFile & "(�汾:" & strSourceVer & "   �޸�����:" & strSourceDate & ")    ====>    " & _
                                    strTargetFile & "(�汾:" & strTargetVer & "   �޸�����:" & strTargetDate & ")        ������Ϣ:" & strErrMsg
                    Exit Function
                End If
            End If
        End If
        
    End With
    Exit Function
ErrHand:
    MsgBox "����ע�Ṥ���ļ�����" & vbCrLf & Err.Description, vbExclamation, "��ʾ"
End Function

Public Sub isMD5UpGrade()

    '--����:����MD5DLL���

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strErrMsg As String
    Dim strTargetFile As String
    Dim strSourceFile  As String
    
    
    On Error GoTo ErrHand:
    strSQL = "Select �޸�����,�汾��,MD5 From zlfilesupgrade where upper(�ļ���)='AAMD532.DLL'"
    Set rsTmp = OpenSQLRecord(strSQL, "���AAMD532")
    
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
                WriteTxtLog strSourceFile & " ====> " & strTargetFile & "������Ϣ:" & strErrMsg
                Exit Sub
            End If
        Else
            If FtpDownFile(strTargetFile, strSourceFile) Then
                WriteTxtLog strSourceFile & " ====> " & strTargetFile & "������Ϣ:" & strErrMsg
                Exit Sub
            End If
        End If
    End With
    
    Exit Sub
ErrHand:
    
End Sub

Public Sub is7zUpGrade()

    '--����:����7Z��ѹ����ѹ���ļ���������

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
    strSQL = "Select �ļ���,�޸�����,�汾��,MD5 From zlfilesupgrade where upper(�ļ���)='7Z.EXE' or upper(�ļ���)='7Z.DLL'"
    Set rsTmp = OpenSQLRecord(strSQL, "���7Z")
    
    With rsTmp
        Do Until rsTmp.EOF
            strSourceFile = gstrWinSystemPath & "\" & NVL(!�ļ���, "")
            If gintUpType = 0 Then
                strTargetFile = gstrServerPath & "\" & NVL(!�ļ���, "")
            Else
                strTargetFile = NVL(!�ļ���, "")
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
             
             '�Ƚ�MD5
             If UCase(strSourceMD5) <> UCase(strTargetMD5) Then
                strTargetVer = GetVersion(IIf(IsNull(!�汾��), 0, !�汾��))
                strTargetDate = Format(!�޸�����, "yyyy-MM-DD HH:mm:ss")
                
                If gintUpType = 0 Then
                    If GetCopyAndReg(strTargetFile, strSourceFile, strErrMsg) Then
    
                        WriteTxtLog strSourceFile & "(�汾:" & strSourceVer & "   �޸�����:" & strSourceDate & ")    ====>    " & _
                                        strTargetFile & "(�汾:" & strTargetVer & "   �޸�����:" & strTargetDate & ")        ������Ϣ:" & strErrMsg
                    End If
                Else
                    If FtpDownFile(strTargetFile, strSourceFile) Then
                        WriteTxtLog strSourceFile & "(�汾:" & strSourceVer & "   �޸�����:" & strSourceDate & ")    ====>    " & _
                                        strTargetFile & "(�汾:" & strTargetVer & "   �޸�����:" & strTargetDate & ")        ������Ϣ:" & strErrMsg
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

    '--����:zlRunas���

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strErrMsg As String
    Dim strTargetFile As String
    Dim strSourceFile  As String
    
    
    On Error GoTo ErrHand:
    strSQL = "Select �޸�����,�汾��,MD5 From zlfilesupgrade where upper(�ļ���)='ZLRUNAS.EXE'"
    Set rsTmp = OpenSQLRecord(strSQL, "���ZLRUNAS")
    
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
                WriteTxtLog strSourceFile & " ====> " & strTargetFile & "������Ϣ:" & strErrMsg
                Exit Sub
            End If
        Else
            If FtpDownFile(strTargetFile, strSourceFile) Then
                WriteTxtLog strSourceFile & " ====> " & strTargetFile & "������Ϣ:" & strErrMsg
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
    '����:��ȡ�������
    '��ȡ�������
    Dim strComputerName As String * 256
    
    On Error Resume Next
    
    Call GetComputerName(strComputerName, 255)
    GetMyCompterName = Trim(Replace(strComputerName, Chr(0), ""))
End Function

Public Sub WriteTxtLog(ByVal strText As String)
    'д��������־
    
    mobjText.WriteLine strText
End Sub

Public Sub OpenLogFile(ByVal blnPreUpgrade As Boolean)
    Dim strTxtFile  As String
    If blnPreUpgrade Then
        strTxtFile = gstrAppPath & "\ZLPerUpGradeList.Lst" 'Ԥ������־
    Else
        strTxtFile = gstrAppPath & "\ZLUpGradeList.Lst" '������־
    End If
    
    '����־�ļ�
    If FindFile(strTxtFile) = False Then
        mobjFile.CreateTextFile (strTxtFile)
    End If
    '--mobjFile.CreateTextFile (strTxtFile)
    
    Set mobjText = mobjFile.OpenTextFile(strTxtFile, ForWriting, True)
End Sub
Public Sub CloseLogFile()
    '�ر���־�ļ�
    mobjText.Close
End Sub

Public Function GetCommpentVersion(ByVal strFile As String) As String

    '����:��ȡָ���ؼ��İ汾��
    '���:
    '����:
    '����:�ɹ�,���ذ汾��,���򷵻ؿ�
    '����:���˺�
    '����:2009-01-16 16:59:34

    Dim objFile As New FileSystemObject
    Dim strVer As String, varVersion As Variant
    
    On Error Resume Next
    '��ȡ�ļ��汾��
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

    '����:��ȡָ���ļ�����ؽ���
    '���:strFile-ָ����DLL�ļ�
    '����:cllOutProcess-���ر����õĽ���ֵ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-01-20 13:59:35

    Dim uProcess As PROCESSENTRY32, uMdlInfor As MODULEENTRY32
    Dim lngMdlProcess As Long, strExeName As String, lngSnapShot As Long, strDLLName As String
    
    On Error GoTo ErrHand:
    '�������̿���
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot > 0 Then
      uProcess.lSize = Len(uProcess)
      If Process32First(lngSnapShot, uProcess) Then
        Do
          '��ý��̵ı�ʶ��
          strExeName = UCase(Left(Trim(uProcess.sExeFile), InStr(1, Trim(uProcess.sExeFile), vbNullChar) - 1))
          If strExeName Like "*" & UCase(strFile) & "*" Then
             'һ����˵ֻ��Exe�ļ��Ż����
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

    '����:����ָ���Ľ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-30 11:06:16

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
'����:��ʼ������ʽ��Ϣ

    On Error GoTo errH
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTemp As String
    
    strSQL = " Select ��Ŀ,���� From zlregInfo where ��Ŀ= '��������'"
    Set rsTmp = OpenSQLRecord(strSQL, "��������")

    If rsTmp.EOF = False Then
        strTemp = NVL(rsTmp!����, "0")
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
'����:��ʼ�ռ���ʽ��Ϣ

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTemp As String
    
    On Error GoTo errH
    strSQL = " Select ��Ŀ,���� From zlregInfo where ��Ŀ= '�ռ���ʽ'"
    Set rsTmp = OpenSQLRecord(strSQL, "�ռ���ʽ")

    If rsTmp.EOF = False Then
        strTemp = NVL(rsTmp!����, "0")
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
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function GetSetupPath(ByVal strFileName As String, ByVal strPathSign As String, ByVal strFileType As String, ByVal strPath As String, ByVal strҵ�񲿼� As String) As String
    '����:��ȡ�ռ��ļ�������·��
    '����:����������·��
    '��� strFileName �ļ�����
    '     strPathSign �ļ���װ·��
    '     strFileType �ļ����� 0�����ļ�,1Ӧ���ļ�,2�����ļ�,3�����ļ�,4�����ؼ�
    '     stPath      ��ǰ���ڵ�Ŀ¼
    '     strҵ�񲿼� �ò������ϲ�ҵ�񲿼�
    '����:ף��
    '����:2010/12/10

    Dim strTemp As String '��ʱ·�����
    Dim strSystemDirectory As String 'ϵͳsystem32Ŀ¼
    Dim strWinDirectory As String  'windowsĿ¼
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
        Case "0" '����
            strTemp = strPath & "\PUBLIC\" & strFileName
        Case "1" 'Ӧ��
            strTemp = strPath & "\Apply\" & strFileName
        Case "2" '����
            strTemp = strWinDirectory & "\Help\" & strFileName
        Case "3" '����
            strTemp = strPath & "\" & strFileName
        Case "4" '����
            strTemp = ""
        Case "5"
            strPathSign = UCase(strPathSign)
            If (InStrRev(strPathSign, "[SYSTEM]", -1) > 0) Or (strPathSign = "") Then
                strTemp = strSystemDirectory & "\" & strFileName
            End If
            
            '��·��
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
            
            '���⴦���������ļ� "ZLAUTORUN.INI"
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
        Else '����·��
            If Left(strFileName, 2) = "\\" Then
                strTemp = ""
            Else
                strTemp = Left(strPath, 1) & Right(strFileName, Len(strFileName) - 1)
            End If
        End If
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''�ж�ҵ�񲿼���׼
'''1.�ȼ��ҵ�񲿼��ڱ����Ƿ��Ѿ�����
'''1.1����
'''    ����
'''1.2������
'''    �������ҵ�񲿼��Ƿ���ҵ���б���
'''    ��
'''        ���ҵ�񲿼��Ƿ���ڣ�ֻҪ��һ�����ڣ�������,�������ء�
'''    ��
'''        ������,˵����������Ҫʹ�����������
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '�ж�Ӧ�ò����Ƿ�װ
    Select Case strFileType
    Case 1 'Ӧ�ò���
        If mobjFile.FileExists(strTemp) Then
            '����
        Else
'      ���Դ���     strҵ�񲿼� = "ZL9BILLEDIT.OCX,ZL9COMLIB.DLL,ZL9DESIGN.OCX,ZL9FUNCTION.DLL,ZL9PRINTMODE.DLL,ZL9REPORT.DLL"
            If Len(strҵ�񲿼�) > 0 Then
                If UCase(GetNoSuffixFilename(strFileName)) = UCase(strҵ�񲿼�) Then GoTo goOn
                strTempProfession = Split(strҵ�񲿼�, ",")
                For i = 0 To UBound(strTempProfession)
                    If UCase(strTempProfession(i)) = UCase("zlSvrStudio") Then '���������⴦��
                        strTempFile = strPath & "\" & strTempProfession(i) & ".exe"
                    Else
                        strTempFile = strPath & "\Apply\" & strTempProfession(i) & ".dll"
                    End If
                    
                    If mobjFile.FileExists(strTempFile) Then
                        blnExits = True
                        Exit For
                    Else
                        If UCase(strFileName) = UCase(strҵ�񲿼�) & ".DLL" Then
                            blnExits = True
                            Exit For
                        End If
                    End If
                Next
                
                If blnExits Then
                    '����
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
    '����:ȷ���Ƿ�Դ����
    '����:��ԭ����-true,����Դ����-false
    On Error Resume Next
    Debug.Print 1 / 0
    IsSourceCode = Err <> 0
End Function

Public Function FileTempDown(ByVal strTargetFile As String, strFileName As String, Optional ByRef strErrMsg As String = "������ɣ�") As Boolean
    Dim strTempDirectored As String '��ʱ�ļ�Ŀ¼
    Dim strPerTempDirectored As String 'Ԥ������ʱ�ļ�Ŀ¼
    Dim strPerFileName As String 'Ԥ�����ļ�
    
    On Error GoTo errH
    If gblnPreUpgrade Then
        strTempDirectored = gstrPerTempPath
    Else
        strTempDirectored = gstrTempPath
    End If
           
    If mobjFile.FolderExists(strTempDirectored) = False Then
       Call mobjFile.CreateFolder(strTempDirectored)
    End If
    
    '��ʱ��ʽ��������
    If gblnOfficialUpgrade Then
        strPerFileName = gstrPerTempPath & strFileName
        If gblnԤ����� Then
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
                    '��ɾ��˳�,��������
                    FileTempDown = True
                    strErrMsg = "������ɣ�"
                    Exit Function
                End If
            End If
        End If
    End If
    
    '��ͨ��������
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
'''            '�ļ�����,�������
'''            If FileSystem.GetAttr(strTargetFile) <> vbNormal Then
'''                    FileSystem.SetAttr strTargetFile, vbNormal
'''            End If
'''        Else
            strErrMsg = "�ļ��ڷ�����Ŀ¼������!"
            FileTempDown = False
            Exit Function
        End If
        
        Err.Clear
        mobjFile.CopyFile strTargetFile, strFileName, True
        If Err <> 0 Then
'            MsgBox "�޷����������б��ļ�" & strTargetFile & vbCrLf & "��ȷ�Ϲ�����������Ƿ����!", vbInformation + vbDefaultButton1, "�ͻ����Զ�����"
            strErrMsg = "�ļ��ڷ�����Ŀ¼������!"
            FileTempDown = False
            Exit Function
        Else
            FileTempDown = True
            strErrMsg = "������ɣ�"
        End If
    Else
        If FtpDownFile(strTargetFile, strFileName) = False Then
'            MsgBox "�޷����������б��ļ�" & strTargetFile & vbCrLf & "��ȷ��FTP���������Ƿ����!", vbInformation + vbDefaultButton1, "�ͻ����Զ�����"
            strErrMsg = "�ļ��ڷ�����Ŀ¼������!"
            FileTempDown = False
            Exit Function
        Else
             FileTempDown = True
             strErrMsg = "������ɣ�"
        End If
    End If
    Exit Function
errH:
    If Err Then
        FileTempDown = False
    End If
End Function

Public Function FileDeCompression(strTempFile As String, Optional ByRef strErrMsg As String = "��ѹ����ɣ�") As Boolean
    '����:��ѹ���ļ�������,��ɾ��ѹ�����ļ�!
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
        strErrMsg = "��ѹ����ɣ�"
        FileDeCompression = True
    Else
        strErrMsg = "��ѹ��ʧ��!"
        FileDeCompression = False
    End If
    Exit Function
errH:
    If Err Then
        FileDeCompression = False
    End If
End Function

Public Function GetErrParameter(ByVal intParameterNum As Integer) As String
'����:����������ȡ����ֵ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "select ����ֵ from ZlOptions where ������=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "������", intParameterNum)

    If rsTmp.RecordCount = 1 Then
         GetErrParameter = NVL(rsTmp!����ֵ, "0")
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
'���ܣ���������Ϣд�����ݿ������־
    Dim strSQL As String
        On Local Error Resume Next
        Dim bytErrType As Byte, lngErrNum As Long
        bytErrType = 4 '�ͻ�����������
        lngErrNum = 0  '�����
        
        
        strSQL = "Insert Into zlErrorLog(�Ự��,�û���,����վ,ʱ��,����,�������,������Ϣ)" & _
            " Select SID,USER,MACHINE,SYSDATE," & bytErrType & "," & lngErrNum & "," & _
            AdjustStr(strMsgInfo) & " From v$Session Where AUDSID=UserENV('SessionID')"
        gcnOracle.Execute strSQL
End Sub

Private Function AdjustStr(Str As String) As String
'���ܣ�������"'"���ŵ��ַ�������ΪOracle����ʶ����ַ�����
'˵�����Զ�(����)�����߼�"'"�綨����

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
'����ļ�·���Ƿ���ڣ��粻���ھʹ���
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
    'ȥ����׺���ļ���
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
    '�ж��Ƿ���й���ԱȨ��,���ɲ����ļ�zlTestAdmin��������system32��
    Dim strAppPath As String
    
    On Error Resume Next
    strAppPath = App.Path & "\zlTestAdmin.txt"
    
    
    Open strAppPath For Output As #1
    Print #1, Now & "   ������ԱȨ��"
    Close #1
    FileCopy strAppPath, GetWinSystemPath & "\zlTestAdmin.txt"

    If Err.Number = 75 Then
        GetAdmin = False
        'MsgBox "û�й���ԱȨ��"
    ElseIf Dir(GetWinSystemPath & "\zlTestAdmin.txt", vbNormal) <> "" Then
        GetAdmin = True
        Call Kill(GetWinSystemPath & "\zlTestAdmin.txt")
        'MsgBox "�й���ԱȨ��"
    Else
        GetAdmin = False
    End If
    
    'ɾ�������ļ�zlTestAdmin
    Call Kill(strAppPath)
End Function

'ϵͳ����Ա������ܷ���
Public Function decipher(stext As String)      '������ܳ���
    Const min_asc = 32 '��СASCII��
    Const max_asc = 126 '���ASCII�� �ַ�
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
       ch = Asc(Mid(stext, i, 1)) 'ȡ��ĸת���ASCII��
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
'���ܣ���������Ϣд�����ݿ������־
    Dim strSQL As String
    On Local Error Resume Next

    strSQL = "Insert Into zltools.zlClientUpdatelog(����վ,��������,����)" & _
        " Select TERMINAL,SYSDATE," & _
        AdjustStr(strMsgInfo) & " From v$Session Where AUDSID=UserENV('SessionID')"
    gcnOracle.Execute strSQL
       
End Sub

Public Sub UpdateCondition(ByVal intMode As Integer)
'���ܣ��������
    Dim strSQL As String
    On Local Error Resume Next

    If intMode = 1 Then
        strSQL = "Zl_Zlclients_Control(15,'" & gstrComputerName & "')"
    Else
        strSQL = "Zl_Zlclients_Control(16,'" & gstrComputerName & "')"
    End If
    Call ExecuteProcedure(strSQL, "UpdateCondition")
End Sub



