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

'Public gobjDemand As Object                '����̨
Public SplashObj As New frmSplash
Public gcnOracle As New ADODB.Connection    '�������ݿ�����

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼

Public gstrUserFlag As String               '��ǰ�û���־(��λ��ʾ)����1λ���Ƿ�DBA����2λ��ϵͳ������

Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public gstrStation As String                '������վ����
Public gstrMenuSys As String                'ϵͳ�˵�

'-----------------------------------------
'�����롢ע���롢�������������ע���������
Public gstrRegCode As String
Public gstrPublish As String
Public gstrParseRegCode As String
Public gstrParsePublish As String
'-----------------------------------------

Public gstrSystems As String

'ȡӲ�̴�С
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
'-ע��� API ����...
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

Private Const STRSPLIT As String = "���"
Private Const REGCMD As String = "REGEDIT /E"
Private Const REGFILE As String = "C:\REGFILE.REG"
Private Const REGDATA As String = "C:\REGDATA.REG"
Private Const REGDIRECTORY As String = """HKEY_CURRENT_USER\SOFTWARE\VB AND VBA PROGRAM SETTINGS\ZLSOFT"""

'---------------------------------------------------------------
'- ע��� Api ����...
'---------------------------------------------------------------
' Reg Data Types...
Const REG_SZ = 1                         ' Unicode���ս��ַ���
Const REG_EXPAND_SZ = 2                  ' Unicode���ս��ַ���
Const REG_DWORD = 4                      ' 32-bit ����

' ע���������ֵ...
Const REG_OPTION_NON_VOLATILE = 0       ' ��ϵͳ��������ʱ���ؼ��ֱ�����

' ע���ؼ��ְ�ȫѡ��...
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
                     
' ע���ؼ��ָ�����...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004

' ����ֵ...
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- ע���ȫ��������...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Public glngOld As Long, glngFormW As Long, glngFormH As Long

'---------------------------------------------------------------
'   ��Ȩ���˵������ð汾
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
    
    'Ϊʵ��XP�������ʾ����ǰ����ִ�иú���
    Call InitCommonControls
    
    BlnShowFlash = False
    Load SplashObj
    '��ע����л�ȡ�û�ע�������Ϣ,����û���λ���Ʋ�Ϊ��,����ʾ���ִ���
    StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��ʾ", "")
    If StrUnitName <> "" Then
        With SplashObj
            '��������Ҫ����
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            Call ApplyOEM_Picture(.imgPic, "PictureB")
            .Show
            .lblGrant = StrUnitName
            StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "������", "")
            If Trim(StrUnitName) = "" Then
                .Label3.Visible = False
                .lbl������.Visible = False
            Else
                .lbl������.Caption = ""
                For IntCount = 0 To UBound(Split(StrUnitName, ";"))
                    .lbl������.Caption = .lbl������.Caption & Split(StrUnitName, ";")(IntCount) & vbCrLf
                Next
            End If
            .LblProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒȫ��", "")
            .lbl����֧���� = GetSetting("ZLSOFT", "ע����Ϣ", "����֧����", "")
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
    
    '�û�ע��
    frmUserLogin.Show 1
    If gcnOracle.State <> adStateOpen Then
        Unload frmUserLogin
        Unload SplashObj
        Exit Sub
    End If
    
    '��ʼ����������
    InitCommon gcnOracle
    If RegCheck = False Then
        Unload SplashObj
        Exit Sub
    End If
    
    '�����������Ч��Ϊ�ջ�Ϊ"-"�������˳�
    gstrParsePublish = zlRegInfo("��Ʒ����")
    gstrParseRegCode = zlRegInfo("��λ����", , -1)
    
    gstrSysName = gstrParsePublish & "���"
    SaveSetting "ZLSOFT", "ע����Ϣ", "��ʾ", gstrSysName
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrSysName"), gstrSysName
    gstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrVersion"), gstrVersion
    gstrAviPath = App.Path & "\�����ļ�"
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrAviPath"), gstrAviPath
    
    With SplashObj
        If BlnShowFlash = False Then
            .lblGrant = gstrParseRegCode
            .lbl����֧����.Caption = zlRegInfo("����֧����", , -1)
            .LblProductName = zlRegInfo("��Ʒ����")
            
            strCode = zlRegInfo("��Ʒ������", , -1)
            .lbl������.Caption = ""
            For IntCount = 0 To UBound(Split(strCode, ";"))
                .lbl������.Caption = .lbl������.Caption & Split(strCode, ";")(IntCount) & vbCrLf
            Next
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            .Show
            BlnShowFlash = True
        End If
        DoEvents
    End With
    
    '���û�ע�������Ϣд��ע���,���´�����ʱ��ʾ
    SaveSetting "ZLSOFT", "ע����Ϣ", "��λ����", gstrParseRegCode
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒȫ��", zlRegInfo("��Ʒ����")
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒ����", zlRegInfo("��Ʒ����")
    SaveSetting "ZLSOFT", "ע����Ϣ", "����֧����", zlRegInfo("����֧����", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "������", zlRegInfo("��Ʒ������", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧���̼���", zlRegInfo("֧���̼���")
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��EMAIL", zlRegInfo("֧����MAIL")
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��URL", zlRegInfo("֧����URL")

    '-------------------------------------------------------------
    '��鱾����װ����
    '-------------------------------------------------------------
    If TestComponent = False Then
        MsgBox "��û�в����κ�ϵͳ��Ȩ�ޣ��������˳���", vbInformation, gstrSysName
        Unload SplashObj
        Exit Sub
    End If
    
    '-------------------------------------------------------------
    '��������ѡ����
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
            gstrSystems = " (ϵͳ in (" & gstrSystems & ") Or ϵͳ Is NULL)"
        End If

        If gstrSystems = "" Then
            MsgBox "��û�в����κ�ϵͳ��Ȩ�ޣ��������˳���", vbInformation, gstrSysName
            Unload SplashObj
            Exit Sub
        End If
        
    End With
    '-------------------------------------------------------------
    '�����˵�������
    '-------------------------------------------------------------
    If StrHaveSys = "" Then
        MsgBox "��û�в����κ�ϵͳ��Ȩ��,�������˳���", vbInformation, gstrSysName
        Unload SplashObj
        Exit Sub
    End If
    
    gstrSQL = "SELECT ϵͳ FROM zlPrograms WHERE ���=1536 AND ϵͳ IN (" & StrHaveSys & ")"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMain")
    
    If gRs.BOF Then
        MsgBox "��û�в����κ�ϵͳ��Ȩ��,�������˳���", vbInformation, gstrSysName
        Unload SplashObj
        Exit Sub
    End If
    
    glngSys = gRs("ϵͳ").Value
    If InStr(1, GetPrivFunc(glngSys, 1536), "����") <= 0 Then
        MsgBox "��û�в����κ�ϵͳ��Ȩ��,�������˳���", vbInformation, gstrSysName
        Unload SplashObj
        Exit Sub
    End If
    
    '-------------------------------------------------------------
    '����ͬ���
    '-------------------------------------------------------------
    Call CreateSynonyms(glngSys, 1536)
    
    Call GetUserInfo
        
    Unload SplashObj
        
    If Not �Ƿ�����ʹ�ñ�����վ Then Exit Sub
        
    Call CodeMan(glngSys, 1536)
    
End Sub

Private Function CreateSynonyms(ByVal lngSys As Long, ByVal lngModul As Long)
    Dim strSQL As String
    '����ģ����������ͬ���(����Ѵ����򲻻��ٴ���)
    On Error Resume Next
    strSQL = "Zl_Createsynonyms(" & lngSys & ")"
    zlDatabase.ExecuteProcedure strSQL, "����ͬ���"
End Function

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
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
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            Else
                MsgBox "�����û�������������ָ�������޷�ע�ᡣ", vbInformation, gstrSysName
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
    '���ܣ� �ر����ݿ�
    '������
    '���أ� �ر����ݿ⣬����True��ʧ�ܣ�����False
    '------------------------------------------------
    Err = 0
    On Error Resume Next
    gcnOracle.Close
    OraDataClose = True
    Err = 0

End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
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
    '���ܣ�����ԱID���޸�������
    '������CurrUser
    '      ��ǰ�û���
    '���أ�����ɹ����˻�True�����򷵻�False
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
'���ܣ�дע���
    Dim rc As Long                                      ' ���ش���
    Dim hKey As Long                                    ' ����һ��ע���ؼ���
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' ע���ȫ����
    
    lpAttr.nLength = 50                                 ' ���ð�ȫ����Ϊȱʡֵ...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- ����/��ע���ؼ���...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' ����/��//KeyRoot//KeyName
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' ������...
    
    '------------------------------------------------------------
    '- ����/�޸Ĺؼ���ֵ...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' Ҫ��RegSetValueEx() ������Ҫ����һ���ո�...
    
    ' ����/�޸Ĺؼ���ֵ
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))
                       
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' ������
    '------------------------------------------------------------
    '- �ر�ע���ؼ���...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' �رչؼ���
    
    UpdateKey = True                                    ' ���سɹ�
    Exit Function                                       ' �˳�
CreateKeyError:
    UpdateKey = False                                   ' ���ô��󷵻ش���
    rc = RegCloseKey(hKey)                              ' ��ͼ�رչؼ���
End Function

'-------------------------------------------------------------------------------------------------
'sample usage - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'-------------------------------------------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'���ܣ���ע���
    Dim i As Long                                           ' ѭ��������
    Dim rc As Long                                          ' ���ش���
    Dim hKey As Long                                        ' ����򿪵�ע���ؼ���
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' ע���ؼ�����������
    Dim tmpVal As String                                    ' ע���ؼ��ֵ���ʱ�洢��
    Dim KeyValSize As Long                                  ' ע���ؼ��ֱ����ߴ�
    
    ' �� KeyRoot {HKEY_LOCAL_MACHINE...} �´�ע���ؼ���
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ��ע���ؼ���
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������...
    
    tmpVal = String$(1024, 0)                             ' ��������ռ�
    KeyValSize = 1024                                       ' ��Ǳ����ߴ�
    
    '------------------------------------------------------------
    ' ����ע���ؼ��ֵ�ֵ...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' ���/�����ؼ��ֵ�ֵ
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ������
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' �����ؼ���ֵ��ת������...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' ������������...
    Case REG_SZ, REG_EXPAND_SZ                              ' �ַ���ע���ؼ�����������
        sKeyVal = tmpVal                                     ' �����ַ�����ֵ
    Case REG_DWORD                                          ' ���ֽ�ע���ؼ�����������
        For i = Len(tmpVal) To 1 Step -1                    ' ת��ÿһλ
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' һ���ַ�һ���ַ�������ֵ��
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' ת�����ֽ�Ϊ�ַ���
    End Select
    
    GetKeyValue = sKeyVal                                   ' ����ֵ
    rc = RegCloseKey(hKey)                                  ' �ر�ע���ؼ���
    Exit Function                                           ' �˳�
    
GetKeyError:    ' ����������������...
    GetKeyValue = vbNullString                              ' ���÷���ֵΪ���ַ���
    rc = RegCloseKey(hKey)                                  ' �ر�ע���ؼ���
End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim RsTmp As New ADODB.Recordset
    
    Set RsTmp = zlDatabase.GetUserInfo
    
    UserInfo.�û��� = gstrDbUser
    UserInfo.���� = gstrDbUser
    If Not RsTmp.EOF Then
        UserInfo.ID = RsTmp!ID
        UserInfo.��� = RsTmp!���
        UserInfo.���� = IIf(IsNull(RsTmp!����), "", RsTmp!����)
        UserInfo.���� = IIf(IsNull(RsTmp!����), "", RsTmp!����)
        UserInfo.�û��� = IIf(IsNull(RsTmp!�û���), "", RsTmp!�û���)
        UserInfo.����ID = IIf(IsNull(RsTmp!����ID), 0, RsTmp!����ID)
        UserInfo.������ = IIf(IsNull(RsTmp!������), "", RsTmp!������)
        UserInfo.���� = IIf(IsNull(RsTmp!������), "", RsTmp!������)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub CodeMan(lngSys As Long, ByVal lngModul As Long)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������

    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
       
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    
    glngSys = lngSys

    gstrUnitName = GetUnitName
'    gblnInsure = (UCase(GetSetting("ZLSOFT", "����ȫ��", "�Ƿ�֧��ҽ��", "")) = UCase("Yes"))
'    gintInsure = Val(GetSetting("ZLSOFT", "����ȫ��", "ҽ�����", 0))
    
    gblnInsure = True
    Call gclsInsure.InitOracle(gcnOracle)
    '-------------------------------------------------
        
    frmMainQuery.Show
    
End Sub

Public Sub InitData()

End Sub

Public Function CloseChildWindows(ByVal frmMain As Object, ByVal FrmSon As Object) As Boolean
    '����:�ر������Ӵ���
    
    Dim FrmThis As Form
    
    On Error Resume Next

    CloseChildWindows = True
    
    For Each FrmThis In Forms
        If FrmThis.Caption <> frmMain.Caption And FrmThis.Caption <> FrmSon.Caption Then Unload FrmThis
    Next
    
    '�رչ��������Ĵ���
    If CloseChildWindows Then CloseChildWindows = CloseWindows

End Function

Public Function TestComponent() As Boolean
    '���û���κβ�����ʹ�ã��򷵻ؼ�
    TestComponent = False
    
    Dim strObjs As String, strCodes As String, strSQL As String
    Dim objComponent As Object
    Dim resComponent As New ADODB.Recordset
    
    '--��ע����ȡ��Ȩ����--
    strObjs = GetSetting("ZLSOFT", "ע����Ϣ", "��������", "")
    If strObjs <> "" Then TestComponent = True: Exit Function
    '--������Ȩ��װ����--
    With resComponent
        strSQL = "Select Distinct Upper(g.����) As ����" & vbCrLf & _
                " From zlPrograms g, zlRegFunc r" & vbCrLf & _
                " Where g.��� = r.��� And Trunc(g.ϵͳ / 100) = r.ϵͳ And Upper(g.����) <> 'ZL9REPORT'"
        
        If .State = adStateOpen Then .Close
        Set resComponent = zlDatabase.OpenSQLRecord(strSQL, "mdlMain")
        Err = 0: On Error Resume Next
        Do While Not .EOF
            Err = 0
            Set objComponent = CreateObject(!���� & ".Cls" & Mid(!����, 4))
            If Err = 0 Then strObjs = strObjs & IIf(strObjs = "", "", ",") & "'" & !���� & "'"
            .MoveNext
        Loop
    End With
    If strObjs = "" Then Exit Function
    TestComponent = True
    SaveSetting "ZLSOFT", "ע����Ϣ", "��������", strObjs

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

Public Function �Ƿ�����ʹ�ñ�����վ() As Boolean
    '��д��:���� 2003-03-09
    '����:�ж��Ƿ�����ù���վʹ�ó��������Ҫ�滻���ز�������ִ���滻�����������Ҫ�������������ǳ��򣬲��ر��˳�
    Dim objFileSys As New FileSystemObject
    Dim rsUse As New ADODB.Recordset
    Dim strSQL As String, strInfo As String, strPath As String, strExeName As String
    Dim blnAllow As Boolean, blnUpdate As Boolean, int������ As Integer
    Dim str�������� As String, Error As Long
    
    On Error Resume Next
    Err = 0
    blnAllow = False
    blnUpdate = False
    str�������� = "zlHisCrust.exe"
    �Ƿ�����ʹ�ñ�����վ = False
    strExeName = GetSetting("ZLSOFT", "����ȫ��", "ִ���ļ�", "")
    
    '�ж��Ƿ�����ʹ��
    strInfo = AnalyseConfigure
    strSQL = "Select Nvl(��ֹʹ��,0) as ����,Nvl(������־,0) as ����,Nvl(�ռ���־,0) as �ռ�,������ From zlClients Where ����վ='" & AnalyseComputer & "'"
    
    rsUse.CursorLocation = adUseClient
    Set rsUse = zlDatabase.OpenSQLRecord(strSQL, "mdlMain")
    With rsUse
        If .EOF Then
            '��û�иù���վ�����ݣ��ϴ���IP����������CPU���ڴ桢Ӳ�̡�����ϵͳ��
            strSQL = " Insert into zlClients" & _
                     " (IP,����վ,CPU,�ڴ�,Ӳ��,����ϵͳ,����)" & _
                     " Values " & _
                     "('" & Split(strInfo, STRSPLIT)(0) & "','" & Split(strInfo, STRSPLIT)(1) & _
                     "','" & Split(strInfo, STRSPLIT)(2) & "','" & Split(strInfo, STRSPLIT)(3) & _
                     "','" & Split(strInfo, STRSPLIT)(4) & "','" & Split(strInfo, STRSPLIT)(5) & _
                     "','" & UserInfo.���� & "')"
            gcnOracle.Execute strSQL
            �Ƿ�����ʹ�ñ�����վ = True
            Exit Function
        Else
            blnAllow = IIf(IIf(IsNull(!����), 0, !����) = 0, True, False)
            int������ = IIf(IsNull(!������), 0, !������) '0-��ʾ������
            blnUpdate = IIf(IIf(IsNull(!����), 0, !����) = 1, True, False)
            If Not blnUpdate Then blnUpdate = (IIf(IsNull(!�ռ�), 0, !�ռ�) = 1)
        End If
    End With
    If Not blnAllow Then
        MsgBox "�ù���վ�ѱ�����Ա���ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�������������
    If int������ > 0 Then
        strSQL = "Select SID From v$Session Where Upper(PROGRAM) Like 'ZLHIS%.EXE' And Status<>'KILLED' And MACHINE=(Select MACHINE From v$Session Where AUDSID=UserENV('SessionID'))"
        If rsUse.State = 1 Then rsUse.Close
        Set rsUse = zlDatabase.OpenSQLRecord(strSQL, "mdlMain")
        If rsUse.RecordCount > int������ Then
            MsgBox "��ǰ����վ���ֻ���� " & int������ & " ����¼���ӣ���ǰ�Ѿ��� " & rsUse.RecordCount - 1 & " �����ӡ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    On Error GoTo errHand
    '���������Ҫ���µı�������������±���ע���
    If Not RegRestoreByManager Then Exit Function
    
    '�����Ҫ�������������ǳ���
    If blnUpdate Then
        On Error Resume Next
        
        strPath = objFileSys.GetParentFolderName(App.Path)
        Error = Shell(strPath & "\" & str�������� & " " & gcnOracle.ConnectionString & "||0||" & strExeName, vbNormalFocus)
        '������ǳ���
        If Error = 0 Then
            MsgBox "δ�ҵ��ͻ����Զ��������ߣ���������ṩ����ϵ��", vbInformation, gstrSysName
            blnAllow = True
        Else
            'MsgBox "����̨���رգ���Ϊ�������ͻ����Զ��������ߣ�", vbInformation, gstrSysName
            blnAllow = False
        End If
    End If
    
    �Ƿ�����ʹ�ñ�����վ = blnAllow
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function AnalyseConfigure() As String
    '��д��:���� 2003-03-09
    '����:���������������ã�IP����������CPU���ڴ桢Ӳ�̡�����ϵͳ��
    Dim strCPU As String           'CPU
    Dim strMemory As String        '�ڴ�
    Dim strOS As String            '����ϵͳ
    Dim strComputerName As String  '�������
    Dim strHD As String            'Ӳ��
    Dim strIP As String            'IP��ַ
    Dim verinfo As OSVERSIONINFO
    Dim sysinfo As SYSTEM_INFO
    Dim memsts As MEMORYSTATUS
    Dim memory&
    
    strIP = AnalyseIP
    
    '��ȡ�������
    strComputerName = AnalyseComputer
    
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
    '����޸�:���˺�,���Ϊ������ʽ����˸����˴˹���
    '��д��:���� 2003-03-08
    '����:�����ݿ��лָ����û�������ע���
    Dim strSection As String, strKey As String, strData As String
    Dim rsReg As New ADODB.Recordset
    Dim rsParaList As New ADODB.Recordset
    Dim strSQL As String, strComputerName As String
    Dim blnUser As Boolean '�Ƿ������û�������
    Dim blnSharedUser As Boolean '�Ƿ���ڲ���վ����û�������У�����Ҫ��Zlclientparaset�в����¼���ָ���־Ϊ2
    On Error GoTo errHand
    RegRestoreByManager = False
    
    strComputerName = AnalyseComputer

    strSQL = "Select ������,����վ,�û���" & _
            " From Zlclientparaset " & _
            " Where ((����վ = '" & strComputerName & "' And �û��� Is Null ) Or  " & _
            "         (����վ Is Null And �û���='" & UCase(gstrDbUser) & "') Or  " & _
            "         (����վ='" & strComputerName & "' And �û���='" & UCase(gstrDbUser) & "')) " & _
            "               And Nvl(�ָ���־, 0) = 1"
    
    zlDatabase.OpenRecordset rsReg, strSQL, "��ȡ�Ƿ���ڲ����ָ�"
    
    If rsReg.RecordCount = 0 Then
        '�������Զ��ָ��Ĳ���
        RegRestoreByManager = True
        Exit Function
    End If
    
    rsReg.Filter = "�û���='" & UCase(gstrDbUser) & "'"
    blnUser = rsReg.RecordCount <> 0
    '�����û������Ƿ��ڱ�վ����������
    blnSharedUser = False
    strSQL = ""
    
    If blnUser Then
        strSQL = "" & _
            "   Select �û��� From Zlclientparaset " & _
            "   where ������=" & Val(zlCommFun.Nvl(rsReg!������)) & _
            "         and ����վ='" & strComputerName & "'  And �û���='" & UCase(gstrDbUser) & "'" & _
            "         And �ָ���־=2"
        zlDatabase.OpenRecordset rsParaList, strSQL, "�����û��Ƿ��ڱ�վ����������"
        If rsParaList.RecordCount <> 0 Then
            blnSharedUser = True
        Else
            strSQL = ",'˽��ȫ��','˽��ģ��'"
        End If
    End If
    rsReg.Filter = "�û���=null and ����վ<> null "
    If rsReg.RecordCount <> 0 Then
        strSQL = " And ��� in ('����ȫ��', '����ģ��'" & strSQL & ")"
        rsReg.MoveFirst
    Else
        If strSQL = "" Then
            Exit Function
        Else
            strSQL = " And ��� in (" & Mid(strSQL, 2) & ")"
        End If
    End If
    rsReg.Filter = 0
    strSQL = "" & _
        "   Select ������,���,���,Ŀ¼,����,��ֵ,������Դ,����˵�� " & _
        "   From zlClientparaList " & _
        "   where ������=" & Val(zlCommFun.Nvl(rsReg!������)) & strSQL
    zlDatabase.OpenRecordset rsParaList, strSQL, "�ָ�����"
    With rsParaList
        Do While Not .EOF
            Select Case zlCommFun.Nvl(!���)
            Case "����ȫ��", "����ģ��"
                Call SaveSetting("ZLSOFT", !��� & IIf(IsNull(!Ŀ¼), "", "\" & !Ŀ¼), !����, IIf(IsNull(!��ֵ), "", !��ֵ))
            Case "˽��ȫ��", "˽��ģ��"
                Call SaveSetting("ZLSOFT", !��� & "\" & UCase(gstrDbUser) & IIf(IsNull(!Ŀ¼), "", "\" & !Ŀ¼), zlCommFun.Nvl(!����), zlCommFun.Nvl(!��ֵ))
            End Select
            .MoveNext
        Loop
    End With
    
    
    '����վ����Ϣ�ָ�
    'zl_zlClientParaSet_Restore(
    '    ������_IN   IN zlClientParaSet.������%type,
    '    ����վ_IN   IN zlClientParaSet.����վ%type,
    '    �û���_IN   IN zlClientParaSet.�û���%TYPE
    strSQL = "zl_zlClientParaSet_Restore("
    strSQL = strSQL & "" & Val(zlCommFun.Nvl(rsReg!������)) & ","
    strSQL = strSQL & "'" & UCase(strComputerName) & "',"
    strSQL = strSQL & "'" & UCase(gstrDbUser) & "')"
    zlDatabase.ExecuteProcedure strSQL, "����ָ���־"
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
    Dim sOS As String
    
    '�����Windows2000�����°汾��������API��ȡһ��
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


