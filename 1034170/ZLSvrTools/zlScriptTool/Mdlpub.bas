Attribute VB_Name = "Mdlpub"
Option Explicit

Public Const HWND_TOPMOST = -1
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000

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
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ע���ؼ��ָ�����...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

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

''''''''��������¼
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
''''''''��������¼

'---------------------------------------------------------------
Public gcnOracle As New ADODB.Connection    '���ݿ�����
Public Const gstrSysName = "�������"              'ϵͳ����
Public gstrUserName As String               '�û���
Public gstr���� As String                   '�û�����
Public gstr��� As String                   '�û����
Public gstrPassword As String               '�û�����
Public gstrServer As String                 '��������
Public gstrSql    As String                 'ͨ�õ�SQL������
Public rsTemp As New ADODB.Recordset        '������ʱrecordset����
Public gstrProductName As String

Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&

Public objFso As New FileSystemObject   '�ļ���������
Public gmobjCommon As Object ' zl9ComLib.clsComLib
Public blnExit As Boolean
Public gmdlOSType As Integer

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, ByRef objConn As ADODB.Connection) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strError As String
    
    On Error GoTo Hwanderr
    If objConn.State = adStateOpen Then objConn.Close
    
    With objConn
         .CursorLocation = adUseClient
         .ConnectionString = "Provider=MSDAORA.1;Password=" & strUserPwd & ";User ID=" & strUserName & ";Data Source=" & strServerName & ";Persist Security Info=True"
         .Open 'MSDAORA.1 Ҫ��ȡORACLE blob�ֶ� Ҫʹ�� Provider= OraOLEDB.Oracle
    End With
    
    OraDataOpen = True
    Exit Function
Hwanderr:
   If err.Number <> 0 Then
        '���������Ϣ
        strError = err.Description
        If InStr(strError, "�Զ�������") > 0 Then
            MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
        ElseIf InStr(strError, "ORA-12154") > 0 Then
            MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
        ElseIf InStr(strError, "ORA-12541") > 0 Then
            MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
        ElseIf InStr(strError, "ORA-01033") > 0 Then
            MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
        ElseIf InStr(strError, "ORA-ORA-03114") > 0 Then
            Exit Function
        Else
            MsgBox "�����û�������������ָ�������޷�ע�ᡣ", vbInformation, gstrSysName
        End If
        OraDataOpen = False
        Exit Function
    End If
    OraDataOpen = False
End Function


Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    'strOld��ԭ����
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


Public Sub SelAll(objTxt As Control)
'���ܣ����ı���ĵ��ı�ѡ��
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
'����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If InStr(strInput, "'") > 0 Or InStr(strInput, """") > 0 Then
        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, gstrSysName
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
'����:����Դ�ļ��е�ָ����Դ���ɴ����ļ�
'����:ID=��Դ��,strExt=Ҫ�����ļ�����չ��(��BMP)
'����:�����ļ���
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
'���ܣ��򿪼�¼��
    If rsTemp.State = adStateOpen Then rsTemp.Close
'    rsTemp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle, adOpenStatic, adLockReadOnly
    If gcnConnect Is Nothing Then Set gcnConnect = gcnOracle
    rsTemp.CursorType = 1
    rsTemp.Open (IIf(strSQL = "", gstrSql, strSQL)), gcnConnect
    '  Call WriteLogFile(IIf(strSQL = "", gstrSql, strSQL))
    End Sub


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


Public Sub ApplyOEM_Picture(objPicture As Object, ByVal str���� As String)
'��Ը���ͼ��Ӧ��OEM����
    Dim strOEM As String
    Dim blnCorp As Boolean
    On Error Resume Next
    
    If gstrProductName <> "-" Then
        '����״̬��ͼ���OEM����
        If gstrProductName <> "����" Then
            If Right(str����, 1) = "B" Then
                '��ʾ��ƷͼƬ
                blnCorp = False
                str���� = Mid(str����, 1, Len(str����) - 1)
            Else
                '��ʾ��˾�ձ�
                blnCorp = True
            End If
            
            strOEM = GetOEM(gstrProductName, blnCorp)
            If str���� = "Picture" Then
                Set objPicture.Picture = LoadCustomPicture(strOEM)
            ElseIf str���� = "Icon" Then
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
    '���ܣ�����ÿ�����ߵ�ASCII��
    '������
    '���أ�
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    'OEMͼƬ���������� ��һ��ָ��˾�ձ꣬��һ���ǲ�Ʒ��ʶ
    strCode = IIf(blnCorp = True, "OEM_", "PIC_")
    For intBit = 1 To Len(strAsk)
        'ȡÿ���ֵ�ASCII��
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function

Public Sub ApplyOEM(objStatus As Object)
'���״̬��Ӧ��OEM����
    Dim strOEM As String
    On Error Resume Next
    
    If gstrProductName <> "-" Then
        objStatus.Panels(1).Text = gstrProductName & "���"
        '����״̬��ͼ���OEM����
        If gstrProductName = "����" Then
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
'���ܣ�ִ��SQL���
    gcnOracle.Execute gstrSql, , adCmdStoredProc
   ' Call WriteLogFile(gstrSql)

End Sub

Public Sub WriteLogFile(ByVal strLogContext As String)
'���ܣ�������������־д����־�ļ�
'������strlogConText ��־����
  Dim filename As String
  Dim logFile As TextStream
  
  filename = App.Path & "\" & Format(Now(), "yyyy-mm-dd") & "_" & "ҩƷ����.log"
  '1���ж���־�ļ��Ƿ��Ѿ�����
  If objFso.FileExists(filename) Then
     Set logFile = objFso.OpenTextFile(filename, ForAppending)
  Else
     Set logFile = objFso.CreateTextFile(filename, True)
  End If
  
  '2��д����־
  logFile.WriteLine (Format(Now(), "yyyy-mm-dd HH:MM:SS") & " �� ")
  logFile.WriteLine (strLogContext)
  logFile.Close
End Sub

Public Function ExcuteSql(ByVal Recordsettest As ADODB.Recordset, ByVal strSQL As String) As Boolean
'���ܣ�16
'������16
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
    MsgBox "�������:" & err.Number & vbCrLf & _
            "��������:" & err.Description, vbCritical + vbOKOnly, "����"
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
     NetR.lpLocalName = "" 'ָ��Ϊ����ӳ���̷� ��X:
     NetR.lpRemoteName = RemoteName '"\\192.168.0.3\�������"
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
' ��������
Dim retOSVersionInf As OSVERSIONINFOEX
Dim retLng As Long

    '�ṹ�ߴ�
    retOSVersionInf.dwOSVersionInfoSize = Len(retOSVersionInf)
    
    '��ȡ Windows �汾
    retLng = GetVersionEx(retOSVersionInf)
    
    If retLng = 0 Then
        GetWindowsVersion = "δ֪"
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
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function


Public Function SearchPrintData(ByVal objVsf As Object, ByRef objPrintVsf As Object, Optional strNotPrintCol As String = "") As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
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
                        objPrintVsf.TextMatrix(lngRow, lngPrintCol) = IIf(Abs(Val(objVsf.TextMatrix(lngRow, lngCol))) = 1, "��", "")
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
    '���ܣ�
    '������
    '���أ�
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
