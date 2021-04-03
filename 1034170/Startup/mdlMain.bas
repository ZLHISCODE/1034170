Attribute VB_Name = "mdlMain"
Option Explicit

Public ZlBrowerDll As Object                '����̨
Public gcnOracle As New ADODB.Connection    '�������ݿ�����
Public gobjRelogin As clsRelogin            '������������Ķ���ʵ��
Public gobjWait As Object                   'չʾ��ģ̬��������ʹ�����˳��Ķ���
Public gstrCommand As String                '����������


Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public gstrUserFlag As String               '��ǰ�û���־(��λ��ʾ)����1λ���Ƿ�DBA����2λ��ϵͳ������

Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public gstrStation As String                '������վ����
Public gstrMenuSys As String                'ϵͳ�˵�

Public gstrSystems As String

Public gobjFile As New FileSystemObject

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'---------------------------------------------------------------
'-ע��� API ����...
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

'�л���ָ�������뷨��
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'��ȡĳ�����뷨������
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'�ж�ĳ�����뷨�Ƿ��������뷨
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

Public Enum REGISTER
    ע����Ϣ
    ˽��ģ��
    ˽��ȫ��
    ����ģ��
    ����ȫ��
End Enum

'---------------------------------------------------------------
'����ʱ�䣬�����ж�������Ļ�ĵȴ�ʱ��
'---------------------------------------------------------------
Public gdtStart As Long

'---------------------------------------------------------------
'   ��Ȩ���˵������ð汾
'---------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'����:������ؽ��̴����API����:2008-10-30 11:34:11:���˺�
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
Public gcll_His_PId As Collection        '�洢��صĽ�����Ϣ:array(��������,PID,���ڸ���),"K"+������

#Const SYS_TRYUSE = "��ʽ" '��ʽ/����


Private Sub SetAppBusyState()
'���������̶���δ�������ʱ���滻��ִ�������̹���ʱ�����ġ����������𡱶Ի���
On Error Resume Next
    App.OleServerBusyMsgTitle = App.ProductName
    App.OleRequestPendingMsgTitle = App.ProductName
    
    App.OleServerBusyMsgText = "���������ڴ����������ĵȴ���"
    App.OleRequestPendingMsgText = "�������������������ĵȴ���"
    
    App.OleServerBusyTimeout = 3000
    App.OleRequestPendingTimeout = 10000
Err.Clear
End Sub

Public Sub Main()
    Dim lngReturn As Long
    Dim StrUnitName As String
    Dim BlnShowFlash As Boolean
    Dim strCode As String, intCount As Integer, strStyle As String, strPath As String
    Dim strTitle As String                  '��Ʒ����
    Dim strTag As String                    '�콢���־
    Dim rsMenu As ADODB.Recordset
    Dim objRIS As Object
    
    gstrCommand = CStr(Command())
    Set gobjRelogin = New clsRelogin
    gobjRelogin.MenuGroup = GetMenuGroup(gstrCommand)
    Call SetAppBusyState
     'Ϊʵ��XP�������ʾ����ǰ����ִ�иú���
    Call InitCommonControls
    BlnShowFlash = False
    If InStr(gstrCommand, "=") <= 0 Then Load frmSplash
    '��ע����л�ȡ�û�ע�������Ϣ,����û���λ���Ʋ�Ϊ��,����ʾ���ִ���
    StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��ʾ", "")
    If StrUnitName <> "" And StrUnitName <> "-" Then
        gdtStart = Timer
        With frmSplash
            '��������Ҫ����
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            Call ApplyOEM_Picture(.imgPic, "PictureB")
            If InStr(gstrCommand, "=") <= 0 Then .Show
            .lblGrant = Replace(StrUnitName, ";", vbCrLf)
            StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "������", "")
            If Trim(StrUnitName) = "" Then
                .Label3.Visible = False
                .lbl������.Visible = False
            Else
                .Label3.Visible = True
                .lbl������.Visible = True
                .lbl������.Caption = ""
                For intCount = 0 To UBound(Split(StrUnitName, ";"))
                    .lbl������.Caption = .lbl������.Caption & Split(StrUnitName, ";")(intCount) & vbCrLf
                Next
            End If
            .LblProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒȫ��", "")
            If Len(.LblProductName) > 10 Then
                .LblProductName.FontSize = 15.75 '����
            Else
                .LblProductName.FontSize = 21.75 '����
            End If
            .lbl����֧���� = GetSetting("ZLSOFT", "ע����Ϣ", "����֧����", "")
            .lbltag = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒϵ��", "")
            
            If Trim$(.lbl����֧����.Caption) = "" Then
                .Label1.Visible = False
                .lbl����֧����.Visible = False
            Else
                .Label1.Visible = True
                .lbl����֧����.Visible = True
            End If
        End With
        Do
            If (Timer - gdtStart) > 1 Then Exit Do
            DoEvents
        Loop
        
        BlnShowFlash = True
        DoEvents
    End If
    '����:14365
    Call zlKillHISPID
    '�û�ע��
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
    
    
    'д�뱾�������������Ϣ
    If IsDesinMode Then
        strPath = GetSetting("ZLSOFT", "����ȫ��", "����·��", "")
        If strPath = "" Then
            strPath = "C:\Appsoft"
        Else
            strPath = Mid(strPath, 1, InStrRev(strPath, "\") - 1)
        End If
        gstrAviPath = strPath & "\�����ļ�"
    Else
        SaveSetting "ZLSOFT", "����ȫ��", "ִ���ļ�", App.EXEName & ".exe"
        SaveSetting "ZLSOFT", "����ȫ��", "����·��", App.Path & "\" & App.EXEName & ".exe"
                
        gstrAviPath = App.Path & "\�����ļ�"
        SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrAviPath"), gstrAviPath
    End If
    
    '2010-05-19 �Զ�������ǰ����Ȩ���ǰִ�С�
    If CheckAllowByTerminal = False Then
        Unload frmSplash
        Exit Sub
    End If
    '��ʼ����������
    InitCommon gcnOracle
    zl9ComLib.SetDbUser gobjRelogin.DBUser
    zl9ComLib.gstrNodeNo = gobjRelogin.NodeNo
    If RegCheck = False Then
        Unload frmSplash
        Exit Sub
    End If
    '�汾���
    Select Case zlRegInfo("��Ȩ����")
        Case "1"
            '��ʽ
            SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", ""
        Case "2"
            '����
            SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", "����"
        Case "3"
            '����
            SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", "����"
        Case Else
            '����
            MsgBox "��Ȩ���ʲ���ȷ���������˳���", vbInformation, gstrSysName
            Unload frmSplash
            Exit Sub
    End Select
    
    gstrSysName = zlRegInfo("��Ʒ����") & "���"
    SaveSetting "ZLSOFT", "ע����Ϣ", "��ʾ", gstrSysName
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrSysName"), gstrSysName
    gstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrVersion"), gstrVersion
    
    strTag = ""
    strTitle = zlRegInfo("��Ʒ����")
    If strTitle <> "" Then
        If InStr(strTitle, "-") > 0 Then
            If Split(strTitle, "-")(1) = "Ultimate" Then
                strTag = "�콢��"
            ElseIf Split(strTitle, "-")(1) = "Professional" Then
                strTag = "רҵ��"
            End If
        End If
    End If
    strTitle = Split(strTitle, "-")(0)
    With frmSplash
        If BlnShowFlash = False Then
            .lblGrant = Replace(zlRegInfo("��λ����", , -1), ";", vbCrLf)
            .lbl����֧����.Caption = zlRegInfo("����֧����", , -1)
            
            .LblProductName = strTitle
            .lbltag = strTag
            strCode = zlRegInfo("��Ʒ������", , -1)
            .lbl������.Caption = ""
            For intCount = 0 To UBound(Split(strCode, ";"))
                .lbl������.Caption = .lbl������.Caption & Split(strCode, ";")(intCount) & vbCrLf
            Next
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            If InStr(gstrCommand, "=") <= 0 Then .Show
            BlnShowFlash = True
        End If
        DoEvents
    End With
    '���û�ע�������Ϣд��ע���,���´�����ʱ��ʾ
    SaveSetting "ZLSOFT", "ע����Ϣ", "��λ����", zlRegInfo("��λ����", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒȫ��", strTitle
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒ����", zlRegInfo("��Ʒ����")
    SaveSetting "ZLSOFT", "ע����Ϣ", "����֧����", zlRegInfo("����֧����", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "������", zlRegInfo("��Ʒ������", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧���̼���", zlRegInfo("֧���̼���")
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��EMAIL", zlRegInfo("֧����MAIL")
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��URL", zlRegInfo("֧����URL")
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒϵ��", strTag
    '��鱾����װ����
    If TestComponent = False Then
        MsgBox "��û�в����κ�ϵͳ��Ȩ�ޣ��������˳���", vbInformation, gstrSysName
        Unload frmSplash
        Exit Sub
    End If
    '��������ѡ����
    With FrmAccoutChoose
        gobjRelogin.Systems = .Show_me
        If .BlnSelect = False Then
            Unload frmSplash
            Exit Sub
        End If
        If gobjRelogin.Systems = "" Then
            MsgBox "��û�в����κ�ϵͳ��Ȩ�ޣ��������˳���", vbInformation, gstrSysName
            Unload frmSplash
            Exit Sub
        End If
    End With
    Call GetUserInfo(IIf(gobjRelogin.Systems = "REPORT", 0, Replace(gobjRelogin.Systems, "'", "")))
    '��¼��Ϣ��ȡ
    gstrDeptName = gobjRelogin.DeptName
    gstrDbUser = gobjRelogin.DBUser
    gstrSystems = gobjRelogin.Systems
    '��ȡ��¼����
    gstrUserFlag = IIf(gobjRelogin.IsSysOwner, "01", "00")
    gstrStation = ComputerName
    If gstrStation = "" Then
        gstrStation = "..."
    End If
    '�����˵�������
    Set rsMenu = MenuGranted(gobjRelogin.MenuGroup)
    If rsMenu.EOF Then
        MsgBox "��û�в����κ�ϵͳ��Ȩ��,�������˳���", vbInformation, gstrSysName
        Unload frmSplash
        Exit Sub
    End If
    '�����ٴ�������ͬ��ʣ��������ڰ�װ������ʱ������˽�е��ڽ���ģ��ʱ����
    'ѡ����ò�ͬ��񵼺�̨
    On Error Resume Next
    Err = 0
    strStyle = zlDatabase.GetPara("����̨", , , "zlBrw")
    Set ZlBrowerDll = CreateObject(strStyle & ".Cls" & Mid(strStyle, 3))
    If Err <> 0 Then
        If strStyle = "ZLBRW" Then
            MsgBox "����ʧ�ܣ������������ļ���ʧ�������°�װ��", vbInformation, gstrSysName
            Unload frmSplash
            Exit Sub
        Else
            Err = 0
            Set ZlBrowerDll = CreateObject("ZLBRW.ClsBrw")
            If Err <> 0 Then
                MsgBox "����ʧ�ܣ������������ļ���ʧ�������°�װ��", vbInformation, gstrSysName
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
    '��������ע������ֵ
    Call UpdateParameters
    Unload frmSplash
    '���������ֹ������ֹ
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
    '���û���κβ�����ʹ�ã��򷵻ؼ�
    TestComponent = False
    
    Dim strObjs As String, strCodes As String, strSql As String
    Dim objComponent As Object
    Dim resComponent As New ADODB.Recordset
    
    On Error GoTo errH
    '--��ע����ȡ��Ȩ����--
    strObjs = GetSetting("ZLSOFT", "ע����Ϣ", "��������", "")
    If strObjs <> "" Then
        If InStr(strObjs, "'ZL9REPORT'") = 0 Then
            If CreateComponent("ZL9REPORT.ClsREPORT") Then
                strObjs = strObjs & ",'ZL9REPORT'"
                SaveSetting "ZLSOFT", "ע����Ϣ", "��������", strObjs
            End If
        End If
        TestComponent = True
        Exit Function
    End If
    '--������Ȩ��װ����--
    strSql = "Select Distinct ���� From (" & _
                " Select Upper(g.����) As ����" & _
                " From zlPrograms g, zlRegFunc r" & _
                " Where g.��� = r.��� And Trunc(g.ϵͳ / 100) = r.ϵͳ" & _
                " Union " & _
                " Select Upper(����) as ���� From zlPrograms Where ��� Between 10000 And 19999)"
    Set resComponent = zlDatabase.OpenSQLRecord(strSql, "")
    With resComponent
        Do While Not .EOF
            If CreateComponent(!���� & ".Cls" & Mid(!����, 4)) Then
                strObjs = strObjs & IIf(strObjs = "", "", ",") & "'" & !���� & "'"
            End If
            .MoveNext
        Loop
    End With
    If strObjs = "" Then Exit Function
    TestComponent = True
    SaveSetting "ZLSOFT", "ע����Ϣ", "��������", strObjs
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
    '���ܣ�������Ȩʹ�ò���װ�Ĳ���������������Ȩʹ�õĲ˵�����
    '������ע����
    '-------------------------------------------------------------
    Dim ArrCommand
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCodes As String
    Dim strObjs As String
    Dim intCount As Integer
    Dim strSystems As String
    Dim BlnOnlySys As Boolean 'ֻ�б���ϵͳ
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
    strObjs = GetSetting("ZLSOFT", "ע����Ϣ", "��������", "")
    If strObjs = "" Then strObjs = "'Zl9Common'"
    strObjs = Replace(strObjs, "','", ",")
    If IsDesinMode Then
        strSql = "Select ���, ID As ���, Nvl(�ϼ�id, 0) As �ϼ�, ����, Decode(Nvl(�̱���,'��'),'��',����,�̱���) as �̱���, ���, ˵��, Nvl(ģ��, 0) As ģ��, Nvl(ϵͳ, 0) As ϵͳ, " & _
                 "        Nvl(ͼ��, 0) As ͼ��, ����, Decode(Upper(RTrim(����)), 'ZL9REPORT', 1, 0) As ���� " & _
                 " From Table(Cast(ZLTOOLS.f_Reg_Menu([1], [2], [3]) As ZLTOOLS.t_Menu_Rowset)) " & _
                 " Union " & _
                 " Select A.���, A.ID, Nvl(�ϼ�id, 0) As �ϼ�, A.����, Decode(Nvl(A.�̱���,'��'),'��',A.����,A.�̱���) As �̱���, A.���, A.˵��, Nvl(A.ģ��, 0) As ģ��, " & _
                 "        Nvl(A.ϵͳ, 0) As ϵͳ, Nvl(ͼ��, 0) As ͼ��, C.����, Decode(C.����, 'ZL9REPORT', 1, 0) As ���� " & _
                 " From (Select Level As ���, ID, �ϼ�id, ����, �̱���, ���, ˵��, Nvl(ģ��,0) ģ��, ϵͳ, ͼ�� " & _
                 "        From zlMenus " & _
                 "        Where ��� = [1] And Nvl(ϵͳ, 0) IN(" & strSYS & ") " & _
                 "        Start With �ϼ�id Is Null " & _
                 "        Connect By Prior ID = �ϼ�id) A, " & _
                 "      (Select ϵͳ, Nvl(ģ��,0) ģ�� " & _
                 "        From zlMenus A " & _
                 "        Where ��� = [1] And Nvl(ϵͳ, 0) IN (" & strSYS & ") " & _
                 "        Minus " & _
                 "        Select ϵͳ * 100, ��� From Zlregfunc Where ϵͳ * 100 IN (" & strSYS & ")) B," & _
                 "      (select ϵͳ, Upper(RTrim(����)) as ����,��� From zlPrograms ) C " & _
                 " Where A.ϵͳ = B.ϵͳ And A.ģ�� = B.ģ�� And A.ģ�� = C.���(+) and A.ϵͳ = C.ϵͳ"

    Else
        strSql = "SELECT ���, Id AS ���, Nvl(�ϼ�id, 0) AS �ϼ�, ����, Decode(Nvl(�̱���,'��'),'��',����,�̱���) As �̱���, ���, ˵��, Nvl(ģ��, 0) AS ģ��, Nvl(ϵͳ, 0) AS ϵͳ, " & _
                 "        Nvl(ͼ��, 0) AS ͼ��, ����, Decode(Upper(Rtrim(����)), 'ZL9REPORT', 1, 0) AS ���� " & _
                 " FROM TABLE(CAST(Zltools.f_Reg_Menu([1], [2], [3]) As " & _
                 " Zltools.t_Menu_Rowset)) "
    End If
    'ʵ�ֱ����������,ģ��ſ�����zlReports.����id,Ҳ������zlRPTGroups.����id,����zlReports
    'ֻ��ȡ��������ģ��ı���
    strSql = "Select ���, ���, �ϼ�, ����, �̱���, ���, ˵��, ģ��, ϵͳ, ͼ��, ����, ����, ������" & vbNewLine & _
                    "From (Select a.*, Decode(a.����, 0, Null, Nvl(b.���, c.���)) ������" & vbNewLine & _
                    "       From (" & strSql & ")  a," & vbNewLine & _
                    "            (Select b.ϵͳ, b.����id, b.���" & vbNewLine & _
                    "              From Zlprograms a, Zlreports b" & vbNewLine & _
                    "              Where Nvl(a.ϵͳ, 0) = Nvl(b.ϵͳ, 0) And a.��� = Nvl(b.����id, 0) And Upper(a.����) = 'ZL9REPORT') b, Zlrptgroups c" & vbNewLine & _
                    "       Where a.ϵͳ = b.ϵͳ(+) And a.ģ�� = b.����id(+) And a.ϵͳ = c.ϵͳ(+) And a.ģ�� = c.����id(+))" & vbNewLine & _
                    "Order By ���, ����, ϵͳ, ģ��, ���, ������"

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
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
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
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "ORA-00604") > 0 Then
                If InStr(strError, "ORA-20002") > 0 Then
                    strError = "��ǰ�û�����ʹ�ø�Ӧ�õ�¼���ݿ⣬����ϵ����Ա��"
                Else
                    strError = "��ǰ�û�����ֹ��¼���ݿ⣬����ϵ����Ա��"
                End If
            End If
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28001") > 0 Then
                MsgBox "�����Ѿ����ڡ�����ϵ����Ա�������룡", vbInformation, gstrSysName
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

Public Function ValEx(ByVal varInput As Variant) As Variant
'���ܣ�����Valֻ�������ֿ�ͷʶ��ValEx�Ե�һ�����ֽ���ʶ��
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

Public Function GetAllSubKey(ByVal KeyRoot As Long, KeyName As String) As Variant
'����:��ȡĳ�����������
'���أ�=��������
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

Public Function ReadStartKey() As String
'���ܣ���ȡע�����������ʼʱ���־(֮һ��Ч����)
    Dim strKey As String
    strKey = GetKeyValue(HKEY_CURRENT_USER, "SOFTWARE\VTCELUS6CS", "IXPHWP")  'FirstStart,1Start
    If strKey = "" Then strKey = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\EG5PZRELSML", "NXPHWP") 'SecondStart,2Start
    If strKey = "" Then strKey = GetKeyValue(HKEY_USERS, ".DEFAULT\SOFTWARE\S1NM9US6CS", "TXPHWP") 'ThirdStart,3Start
    If strKey <> "" Then ReadStartKey = CStr(CDate(strKey))
End Function

Public Function WriteStartKey() As Boolean
'����:��ע�����д������ʼʱ���־
    Dim curDate As Date
    curDate = Format(Date, "yyyy-MM-dd")
    WriteStartKey = UpdateKey(HKEY_CURRENT_USER, "SOFTWARE\VTCELUS6CS", "IXPHWP", CCur(curDate)) 'FirstStart,1Start
    WriteStartKey = WriteStartKey And UpdateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\EG5PZRELSML", "NXPHWP", CCur(curDate)) 'SecondStart,2Start
    WriteStartKey = WriteStartKey And UpdateKey(HKEY_USERS, ".DEFAULT\SOFTWARE\S1NM9US6CS", "TXPHWP", CCur(curDate)) 'ThirdStart,3Start
End Function

Public Function ReadValidKey() As String
'���ܣ���ȡע������������ڱ�־(֮һ��Ч����)
    Dim strKey As String
    strKey = GetKeyValue(HKEY_CURRENT_USER, "SOFTWARE\PZ7Q64F9", "IRSUTR") 'OneValid,1Valid
    If strKey = "" Then strKey = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\SDDQ64F9", "NRSUTR") 'TwoValid,2Valid
    If strKey = "" Then strKey = GetKeyValue(HKEY_USERS, ".DEFAULT\SOFTWARE\S1CKGZHPNO", "TRSUTR") 'ThreeValid,3Valid
    If strKey <> "" Then ReadValidKey = strKey
End Function

Public Function WriteValidKey() As Boolean
    '����:��ע�����д�������ڱ�־
    WriteValidKey = UpdateKey(HKEY_CURRENT_USER, "SOFTWARE\PZ7Q64F9", "IRSUTR", "Q64F9") 'OneValid,1Valid
    WriteValidKey = WriteStartKey And UpdateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\SDDQ64F9", "NRSUTR", "Q64F9") 'TwoValid,2Valid
    WriteValidKey = WriteStartKey And UpdateKey(HKEY_USERS, ".DEFAULT\SOFTWARE\S1CKGZHPNO", "TRSUTR", "Q64F9") 'ThreeValid,3Valid
End Function

Public Function GetUserInfo(ByVal strSystems As String)
    Dim rsTmp As New ADODB.Recordset, rsUser As New ADODB.Recordset
    Dim strSql As String, i As Integer
    '���û���Ϣ���蹫����������������ʹ��
    
    With rsTmp
        If .State = adStateOpen Then .Close
        strSql = "Select S.*" & _
                " From zlSystems S,(Select Distinct owner From All_Tables Where Table_Name='���ű�') D" & _
                " Where Upper(S.������)=D.Owner And S.��� In (" & strSystems & ") Order by S.���"
        .Open strSql, gcnOracle, adOpenKeyset
        If Not .EOF Then
            '��Ϊ���ܸ��û����ж��ϵͳ����ݣ�����ѭ��ȡ���
            glngUserId = 0 '��ǰ�û�id
            gstrUserCode = "" '��ǰ�û�����
            gstrUserName = "" '��ǰ�û�����
            gstrUserAbbr = "" '��ǰ�û�����
            glngDeptId = 0 '��ǰ�û�����id
            gstrDeptCode = "" '��ǰ�û�
            gstrDeptName = "" '��ǰ�û�
            
            For i = 1 To .RecordCount
                strSql = "Select R.*,D.���� as ���ű���,D.���� as ��������,P.���,P.����,P.����" & _
                        " From " & !������ & ".�ϻ���Ա�� U," & !������ & ".��Ա�� P," & !������ & ".���ű� D," & !������ & ".������Ա R" & _
                        " Where U.��ԱID = P.ID And R.����ID = D.ID And P.ID=R.��ԱID and U.�û���=USER And (P.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.����ʱ�� Is Null) and R.ȱʡ=1"
                Set rsUser = New ADODB.Recordset
                rsUser.CursorLocation = adUseClient
                rsUser.Open strSql, gcnOracle, adOpenKeyset
                Set rsUser.ActiveConnection = Nothing
                If Not rsUser.EOF Then
                    glngUserId = rsUser!��ԱID '��ǰ�û�id
                    gstrUserCode = rsUser!��� '��ǰ�û�����
                    gstrUserName = IIf(IsNull(rsUser!����), "", rsUser!����) '��ǰ�û�����
                    gstrUserAbbr = IIf(IsNull(rsUser!����), "", rsUser!����) '��ǰ�û�����
                    glngDeptId = rsUser!����ID '��ǰ�û�����id
                    gstrDeptCode = rsUser!���ű��� '��ǰ�û�
                    gstrDeptName = rsUser!�������� '��ǰ�û�
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
    '--����Ƿ�Դ���뻷��
    If App.EXEName = "prjMain" Then RunningInIDE = True
End Function

'**********************************************************************************************************************
'����:���´�����ؽ��̵ĺ���
'����:���˺�
'����:2008-10-30 11:38:58
Public Function zlKillHISPID() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:ɱ������HIS����������Ƴ�����(ɱ��������:����ZLHIS+.exe�Ľ��������κδ���)
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-30 11:06:16
    '-----------------------------------------------------------------------------------------------------------
    Dim lngProcess As Long, i As Long
    
    zlKillHISPID = False
    Err = 0: On Error GoTo Errhand:
    '��һ��:��Ҫ������ص�ZLHIS����ؽ���
    Set gcll_His_PId = New Collection
    If zlHISPidToCollect(gcll_His_PId) = False Then zlKillHISPID = True: Exit Function  '���������صĴ��󣬾�ֱ�ӷ�����
    If gcll_His_PId Is Nothing Then zlKillHISPID = True: Exit Function
    If gcll_His_PId.Count = 0 Then zlKillHISPID = True: Exit Function
    
    '�ڶ���:��Ҫ�������ZLHIS����ؽ��̵���ش��ڸ���,�����ź��жϳ���صĽ����Ƿ�����쳣,�����쳣�ģ��͵�ɱ��
    Call EnumWindows(AddressOf EnumWindowsProc, 0&)
    For i = 1 To gcll_His_PId.Count
        If Val(gcll_His_PId(i)(2)) <= 1 Then
            '�϶�������С��1����,��ô�϶����쳣����Ҫɱ����
            If Val(gcll_His_PId(i)(1)) <> 0 Then
                '����δ�ɹ������޴���������
                Call TerminatePID(Val(gcll_His_PId(i)(1)))
            End If
        End If
    Next
    zlKillHISPID = True
Errhand:
End Function

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ���д��ڷ���HIS�Ľ��̵Ĵ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-30 10:26:02
    '-----------------------------------------------------------------------------------------------------------
    Dim strTittle As String, lngPID As Long, strName As String
    Dim lngCount As Long
    
    If GetParent(hwnd) = 0 Then
        '��ȡ hWnd ���Ӵ�����
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
    EnumWindowsProc = True ' ��ʾ�����о� hWnd
    Exit Function
End Function

Private Function TerminatePID(ByVal lngPID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ָ���Ľ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-30 11:06:16
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
    '����:��ȡZLHIS�Ľ��̸���صļ���(gcll_HIS_Pid)
    '���:
    '����:cll_His_Pid-������HIS.exe�ĳ���װ�ظü�����
    '����:
    '����:���˺�
    '����:2008-10-30 10:07:38
    '-----------------------------------------------------------------------------------------------------------
    Dim strEXEName  As String, lngSnapShot As Long, lngProcess As Long, lngCount  As Long
    Dim strCurExeName As String, lngCurPid As Long
    Dim uProcess   As PROCESSENTRY32
    Dim StrSessionID As String '��ǰ�ỰID
    Dim StrHISSessionID As String '����ZLHIS���̻ỰID
    Const TH32CS_SNAPPROCESS = &H2
    
    
    Err = 0: On Error GoTo Errhand:
    strCurExeName = "*" & UCase(App.EXEName) & "*"
    
    lngCurPid = GetCurrentProcessId '��ȡ��ǰӦ�ó������
    lngSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    
    StrSessionID = GetCurSessionID(lngCurPid)
    
    
    If lngSnapShot <> 0 Then
        uProcess.lSize = Len(uProcess)
        lngProcess = ProcessFirst(lngSnapShot, uProcess)
        lngCount = 0
        Do While lngProcess
            '�����ڵ�ǰ���̵ĲŴ���
            If lngCurPid <> uProcess.lProcessId Then
                strEXEName = UCase(Left(uProcess.sExeFile, InStr(1, uProcess.sExeFile, vbNullChar) - 1))
                If strEXEName Like strCurExeName Then '"ZLHIS+.EXE"
                    StrHISSessionID = GetCurSessionID(uProcess.lProcessId)
                    '�����ǰzlhis+�Ľ��̻ỰID�������ĻỰID��ͬ,�Ž��йرմ���
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
    '����:��ȡ��ǰ���̵ĻỰID
    '���:��ǰ����PID
    '����:
    '����:�ỰID
    '����:ף��
    '����:2012-06-06 10:15:00
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
    '���ܣ��Ƿ���64λϵͳ
    '���أ�
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
    '--����Ȩ�޲˵�--
    If strCommand = "" Then
        GetMenuGroup = "ȱʡ"
    Else
        ArrCommand = Split(gstrCommand, " ")
        If UBound(ArrCommand) = 0 Then
            '���������˵�����������/����ʾ���û�������ĸ�ʽ���磺zlhis/his��
            If InStr(1, ArrCommand(0), "/") = 0 And InStr(ArrCommand(0), ",") = 0 Then
                GetMenuGroup = ArrCommand(0)
            Else
                GetMenuGroup = "ȱʡ"
            End If
        Else
            '�û��������뼰�˵����
            If UBound(ArrCommand) = 2 And InStr(ArrCommand(0), "=") <= 0 Then
                GetMenuGroup = ArrCommand(2)
            Else
                GetMenuGroup = "ȱʡ"
            End If
        End If
    End If
End Function

