Attribute VB_Name = "mdlMain"
Option Explicit
Public gstrDBUser As String
Public gcnOracle As ADODB.Connection
Public gblnOwner As Boolean
Public gstrSysname As String '��������

Public gstrSystems As String 'ϵͳ����
Public gstr�û���λ���� As String '�ѵ�¼ʱ��Ϊ��

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
Const REG_SZ = 1                         ' Unicode���ս��ַ���
Const REG_EXPAND_SZ = 2                  ' Unicode���ս��ַ���
Const REG_DWORD = 4                      ' 32-bit ����

Public Const WinStyle = &H40000

' ע���ؼ��ְ�ȫѡ��...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ע���ؼ��ָ�����...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_USERS = &H80000003

' ����ֵ...
Public Const ERROR_SUCCESS = 0
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

'---��дINI�ļ���API����
#If Win32 Then
   Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal Filename As String) As Integer
#Else
   Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal Filename As String) As Integer
#End If
'----------------------

Public Enum �����嵥
    ���������嵥 = 10
    �ֵ������ = 11
    ��Ϣ�շ����� = 12
    ϵͳѡ������ = 13
    EXCEL������ = 14
    ���ز������� = 15
End Enum

Public Sub Main()
    
    Call InitCommonControls
    
    gblnOwner = False
    gstrDBUser = ""
    gstrSysname = "����˵���Ķ���"
    gstr�û���λ���� = ""
    
    '�û�ע��
    frmUserLogin.Show 1
    If gcnOracle Is Nothing Then
        Set gcnOracle = New ADODB.Connection
    End If
    
    If gcnOracle.State = adStateOpen Then
        '��ʼ����������
        InitCommon gcnOracle
        
        If RegCheck = False Then
            Exit Sub
        End If
        
        '-------------------------------------------------------------
        '�汾���
        '-------------------------------------------------------------
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
                MsgBox "��Ȩ���ʲ���ȷ���������˳���", vbInformation, gstrSysname
                Exit Sub
        End Select
    
        '�����ס�ZYB��2001-09-19�޸�
        '-------------------------------------------------------------
        '��鱾����װ����
        '-------------------------------------------------------------
        If TestComponent = False Then
            MsgBox "��û�в����κ�ϵͳ��Ȩ�ޣ��������˳���", vbInformation, gstrSysname
            Exit Sub
        End If
        
        '-------------------------------------------------------------
        '��������ѡ����
        '-------------------------------------------------------------
        With FrmAccoutChoose
            gstrSystems = .Show_me
            If .BlnSelect = False Then
                Exit Sub
            End If

            If gstrSystems = "" Then
                MsgBox "��û�в����κ�ϵͳ��Ȩ�ޣ��������˳���", vbInformation, gstrSysname
                Exit Sub
            End If
            
            If gstrSystems <> "REPORT" Then
                gstrSystems = " ϵͳ in (" & gstrSystems & ")"
            End If
        End With
        
        '-------------------------------------------------------------
        '�����˵�������
        '-------------------------------------------------------------
        
        Set rsMenu = MenuGranted("")
        Set rsMenuPEIS = MenuGranted("PEIS")
        
        If rsMenu.EOF Then
            MsgBox "��û�в����κ�ϵͳ��Ȩ��,�������˳���", vbInformation, gstrSysname
            Exit Sub
        End If
        
        gstr�û���λ���� = zlRegInfo("��λ����", , -1)
        
        Call frmMain.Show_me(1) '0- δ��¼��ʽ 1���ѵ�¼��ʽ
    Else
        Call frmMain.Show_me(0) '0- δ��¼��ʽ 1���ѵ�¼��ʽ
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
            MsgBox "����ʧ�ܣ�����ȷ���û�����������������", vbInformation, App.Title
            Err.Clear: Exit Function
        End If
    End With
        
    '�Ƿ��������û�
    If UCase(strUserName) <> "SYS" And UCase(strUserName) <> "SYSTEM" Then
        strSql = "Select 1 From zlSystems Where ������=USER"
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


Private Function TestComponent() As Boolean
    '���û���κβ�����ʹ�ã��򷵻ؼ�
    TestComponent = False
    
    Dim strObjs As String, strCodes As String, strSql As String
    Dim objComponent As Object
    Dim resComponent As New ADODB.Recordset
    
    '--��ע����ȡ��Ȩ����--
    strObjs = GetSetting("ZLSOFT", "ע����Ϣ", "��������", "")
    If strObjs <> "" Then TestComponent = True: Exit Function
    '--������Ȩ��װ����--
    With resComponent
        strSql = "Select Distinct Upper(g.����) As ����" & vbCrLf & _
                " From zlPrograms g, zlRegFunc r" & vbCrLf & _
                " Where g.��� = r.��� And Trunc(g.ϵͳ / 100) = r.ϵͳ And Upper(g.����) <> 'ZL9REPORT'"
        
        If .State = adStateOpen Then .Close
        .Open strSql, gcnOracle, adOpenKeyset
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

Private Function MenuGranted(ByVal strMenuGroup As String) As ADODB.Recordset
    '-------------------------------------------------------------
    '���ܣ�������Ȩʹ�ò���װ�Ĳ���������������Ȩʹ�õĲ˵�����
    '������ע����
    '-------------------------------------------------------------
    Dim ArrCommand
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCodes As String
    Dim strObjs As String
    Dim IntCount As Integer
    Dim strSystems As String
    Dim gstrMenuSys As String
    Dim BlnOnlySys As Boolean 'ֻ�б���ϵͳ
    Dim strSYS As String
    
    BlnOnlySys = (gstrSystems = "REPORT")
    If BlnOnlySys Then
        strSystems = " '0'"
        strSYS = strSystems
    Else
        strSystems = Replace(gstrSystems, "ϵͳ in (", "")
        strSystems = Replace(strSystems, ")", "")
        strSYS = strSystems
        strSystems = Replace(strSystems, "','", ",")
    End If
    
    '--����Ȩ�޲˵�--
    With rsTemp
        If Command() = "" Then
            gstrMenuSys = "ȱʡ"
        Else
            ArrCommand = Split(Command(), " ")
            If UBound(ArrCommand) = 0 Then
                '���������˵�����������/����ʾ���û�������ĸ�ʽ���磺zlhis/his��
                If InStr(1, ArrCommand(0), "/") = 0 Then
                    gstrMenuSys = ArrCommand(0)
                Else
                    gstrMenuSys = "ȱʡ"
                End If
            Else
                '�û��������뼰�˵����
                If UBound(ArrCommand) = 2 Then
                    gstrMenuSys = ArrCommand(2)
                Else
                    gstrMenuSys = "ȱʡ"
                End If
            End If
        End If
        If strMenuGroup <> "" Then gstrMenuSys = strMenuGroup
        strObjs = GetSetting("ZLSOFT", "ע����Ϣ", "��������", "")
        If strObjs = "" Then strObjs = "'Zl9Common'"
        strObjs = Replace(strObjs, "','", ",")

        strSql = "SELECT ���, Id AS ���, Nvl(�ϼ�id, 0) AS �ϼ�, ����, Decode(Nvl(�̱���,'��'),'��',����,�̱���) As �̱���, ���, ˵��, Nvl(ģ��, 0) AS ģ��, Nvl(ϵͳ, 0) AS ϵͳ, " & _
                 "        Nvl(ͼ��, 0) AS ͼ��, nvl(����,'0') as ����, Decode(Upper(Rtrim(����)), 'ZL9REPORT', 1, 0) AS ���� " & _
                 " FROM TABLE(CAST(Zltools.f_Reg_Menu('" & gstrMenuSys & "', " & strSystems & ", " & strObjs & ") As " & _
                 " Zltools.t_Menu_Rowset)) " & _
                 " ORDER BY ���, Id"

        If .State = adStateOpen Then .Close
        .Open strSql, gcnOracle, adOpenKeyset
    End With
    
    Set MenuGranted = rsTemp
    
End Function

Public Sub WriteToIni(ByVal Filename As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
''дINI�ļ�
    Dim buff As String * 128
    buff = Trim(Value) + Chr(0)
    WritePrivateProfileString Section, Key, buff, Filename

End Sub

Public Function ReadFromIni(ByVal Filename As String, ByVal Section As String, ByVal Key As String) As String
''��INI�ļ�
    Dim i As Long
    Dim buff As String * 128
    GetPrivateProfileString Section, Key, "", buff, 128, Filename
    i = InStr(buff, Chr(0))
    ReadFromIni = Trim(Left(buff, i - 1))
End Function
