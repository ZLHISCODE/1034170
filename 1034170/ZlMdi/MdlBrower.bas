Attribute VB_Name = "MdlBrower"
Option Explicit
'MDI����
Public Type Menu_Type
    ���ܲ˵� As Long
    ���ڲ˵� As Long
    �������ܲ˵� As Long
    �ָ��˵� As Long
End Type
Public �˵���׼ As Menu_Type
Public Enum �����嵥
    ���������嵥 = 10
    �ֵ������ = 11
    ��Ϣ�շ����� = 12
    ϵͳѡ������ = 13
    EXCEL������ = 14
    ���ز������� = 15
End Enum
'��ҹ���
Public gobjPlugIn As Object

Public gobjRelogin As Object                   '���������
Public FrmMainface As Form
Public gcnOracle As ADODB.Connection

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

Public gstrObj() As String
Public gobjCls() As Object
Public grsMenus As New ADODB.Recordset       '�˵���¼��
Public gstrMenuSys As String                '�˵�����
Public gstrCommand As String                '�����в��� �¶� 2010-12-06
Private mlngSysPre As Long                  '�ϴε���˽��ͬ��ʼ�鴴��ʱ��ϵͳ��
Private mlngWin32 As Long
Private mblnע�� As Boolean

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const Process_Query_Information = &H400
Private Const Still_Active = &H103
'---------------------------------------------------------------------------------------------------
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'�ر�ϵͳ��صı�����API����
'----------------------------------------------------------------------------------------------------
Public Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type
Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetVersion Lib "kernel32" () As Long
'The GetCurrentProcess function returns a pseudohandle for the current process.
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
'The OpenProcessToken function opens the access token associated with a process.
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
'The LookupPrivilegeValue function retrieves the locally unique identifier (LUID) used on a specified system to locally represent the specified privilege name.
Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
'The AdjustTokenPrivileges function enables or disables privileges in the specified access token. Enabling or disabling privileges in an access token requires TOKEN_ADJUST_PRIVILEGES access.
Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Public Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Public Declare Function GetLastError Lib "kernel32" () As Long
'����ExitWindowsEx
Private Const M_lng�رռ��������Դ As Long = 8
Public Const EWX_FORCE = 4 'ǿ�йرճ���ע��
'�Զ���
Public Const WINDOWS95 = 0
Public Const WINDOWSNT = 1

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer

Public Sub ExecuteFunc(lngSys As Long, Components As String, Modul As Long, Optional ByVal strPara As String) ', Identity As Byte
    '-------------------------------------------------------------
    '���ܣ�����ִ��ָ�������Ĺ��ܳ���
    '������
    '   frmbrower��������
    '   Components������
    '   Modul��ģ����
    '   Identity����ִ�������Ҫ��
    '-------------------------------------------------------------
    Dim rsCheck As New ADODB.Recordset                  '���汾�Ƿ����ϵͳ����
    Dim IntCount As Integer, intClients As Integer
    Dim objNow As Object                                '�����Ĳ�������
    Dim BlnExecute As Boolean                           '�Ƿ���ڸò���
    Dim StrVersion As String, StrCompareVersion As String
    Dim ArrayVersion
    '�Ϸ��Լ��
    Dim intAtom As Integer, strCommon As String
    Dim strSQL  As String
    
    Err = 0: On Error Resume Next
    FrmMainface.MousePointer = 11
    
    IntCount = UBound(gstrObj)
    If Err <> 0 Then IntCount = -1
    Err = 0
    
    BlnExecute = False
    If IntCount >= 0 Then
        For IntCount = 0 To UBound(gstrObj)
            If gstrObj(IntCount) = Components Then
                BlnExecute = True
                Exit For
            End If
        Next
    End If
    
    'ʹ���²�������
    If UCase(Components) = UCase("zl9EmrInterface") And BlnExecute = False Then
        IntCount = UBound(gstrObj)
        IntCount = IntCount + 1
        ReDim Preserve gstrObj(IntCount)
        gstrObj(IntCount) = Components
        If FrmMainface.mobjEmr Is Nothing Then
            MsgBox "�����������ʧ�ܣ����鲢���µ�¼��", vbInformation, gstrSysName
            Exit Sub
        ElseIf FrmMainface.mobjEmr.IsInited = False Then
            MsgBox "�������δ�ܳ�ʼ��," & FrmMainface.mobjEmr.GetError(), vbInformation, gstrSysName
            Exit Sub
        End If
        Dim strSpecify As String 'Ƭ�Σ�����Ȩ�޹̶��ڵ���ǰ����
        If Not FrmMainface.mobjEmr.HasInjectAuthorization(2201) Then
            strSpecify = GetPrivFunc(lngSys, 2201)
            Call FrmMainface.mobjEmr.InjectAuthorization(2201, strSpecify)
        End If
        If Not FrmMainface.mobjEmr.HasInjectAuthorization(2203) Then
            strSpecify = GetPrivFunc(lngSys, 2203)
            Call FrmMainface.mobjEmr.InjectAuthorization(2203, strSpecify)
        End If
        BlnExecute = True
    End If
    '--���û�иò���,�򴴽�--
    If BlnExecute = False Then
        Set objNow = CreateObject(Components & ".Cls" & Mid(Components, 4))
    
        If Err = 0 Then
            On Error GoTo errH
            '--���ò����İ汾�Ƿ�����ϵͳ����(���汾-3;�ΰ汾-3;���汾-3)[�Զ��屨��������]--
            If Not (UCase(Components) = "ZL9REPORT") And Not (UCase(Components) = "ZL9DOC") And Not OS.IsDesinMode Then
                strSQL = " Select nvl(���汾,1) ���汾,nvl(�ΰ汾,0) �ΰ汾,nvl(���汾,0) ���汾,���� " & _
                          " From ZlComponent Where Upper(Rtrim(����))=[1] And ϵͳ=[2]"
                Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "�����汾���", UCase(Components), lngSys)
                With rsCheck
                    If .EOF Then
                        MsgBox "ϵͳ������ZlComponent���ݲ����������������������ϵ��", vbInformation, gstrSysName
                        FrmMainface.MousePointer = 0
                        Exit Sub
                    End If
                    
                    '��װ�汾��Ϊ��λ���汾����λ�ΰ汾����λ���汾
                    StrCompareVersion = String(3 - Len(!���汾), "0") & !���汾 & "." & _
                                        String(3 - Len(!�ΰ汾), "0") & !�ΰ汾 & "." & _
                                        String(3 - Len(!���汾), "0") & !���汾
                    ArrayVersion = Split(objNow.Version, ".")
                    StrVersion = String(3 - Len(ArrayVersion(0)), "0") & ArrayVersion(0) & "." & _
                                 String(3 - Len(ArrayVersion(1)), "0") & ArrayVersion(1) & "." & _
                                 String(3 - Len(ArrayVersion(2)), "0") & ArrayVersion(2)
                    
                    If StrVersion < StrCompareVersion Then
                        MsgBox "�ò����İ汾�Ѳ�������ϵͳ���������������������ϵ����" & !���� & "��", vbInformation, gstrSysName
                        FrmMainface.MousePointer = 0
                        Exit Sub
                    End If
                End With
            End If
        
            IntCount = 0
            On Error Resume Next
            IntCount = UBound(gstrObj)
            IntCount = IntCount + 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo errH
            ReDim Preserve gobjCls(IntCount)
            Set gobjCls(IntCount) = objNow
            ReDim Preserve gstrObj(IntCount)
            gstrObj(IntCount) = Components
        '��������ʧ�ܣ�Ӧ����ʾ
        Else
            Screen.MousePointer = 0
            MsgBox "���� " & Components & ".Cls" & Mid(Components, 4) & " �����������������鰲װ�Ƿ���ȷ����Ϣ��" & vbNewLine & Err.Description, vbExclamation, gstrSysName
            Err.Clear
            Exit Sub
        End If
    End If
    
    Err = 0: On Error GoTo errH
    '--ִ�иù���--
    If gstrObj(IntCount) = Components Then
        If UCase(Components) = "ZL9REPORT" Then
            If Modul = �˵���׼.�������ܲ˵� Then
                gobjCls(IntCount).ReportMan gcnOracle, FrmMainface
            Else
                
'                strPara = "��ʼ����=2013-01-01"
                If strPara <> "" Then
                    Dim varPara As Variant
                                        
                    varPara = Split(strPara, "|")
'                    varPara(0) = "��ʼ����=2013-01-01"
'                    varPara(1) = "��������=2014-05-01"
                    
                    '���֧��10������������10���Ĳ���
                    Select Case UBound(varPara)
                    Case 0
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0))
                    Case 1
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1))
                    Case 2
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2))
                    Case 3
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3))
                    Case 4
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4))
                    Case 5
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5))
                    Case 6
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6))
                    Case 7
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7))
                    Case 8
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7)), CStr(varPara(8))
                    Case 9
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7)), CStr(varPara(8)), CStr(varPara(9))
                    Case Else
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7)), CStr(varPara(8)), CStr(varPara(9))
                    End Select
                    
                Else
                    gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface
                End If
                
            End If
        ElseIf UCase(Components) = UCase("zl9EmrInterface") Then
            Dim strFuncs As String, strModul As String
            
            strSQL = " Select ���⡡From zlPrograms Where ���=[1] "
            Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "ϵͳģ����", Modul)
            With rsCheck
                    If .EOF Then
                        MsgBox "ϵͳ�����ݲ����������������������ϵ��", vbInformation, gstrSysName
                        FrmMainface.MousePointer = 0
                        Exit Sub
                    Else
                        strModul = !����
                    End If
            End With
            strFuncs = GetPrivFunc(lngSys, Modul)
            Call FrmMainface.mobjEmr.CodeMain(Modul, strModul, FrmMainface.hwnd, gobjRelogin.InputUser, IIf(gobjRelogin.IsTransPwd, "", "[DBPASSWORD]") & gobjRelogin.InputPwd, strFuncs)
        Else
            Call CreateSynonyms(lngSys, Modul)
            
            '�û�վ������� (��ʽ�漰���ð�)
            intClients = Val(zlRegInfo("��Ȩվ��"))
            If intClients > 0 Then
                If GetCurStates > intClients Then
                    MsgBox "��ǰ�û���¼�������������Ȩ��" & intClients & ",ϵͳ���Զ��������У�", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If

            
            'ΪͨѶԭ�Ӹ�ֵ
            strCommon = Format(Now, "yyyyMMddHHmm")
            strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
            '����ͨѶԭ��
            intAtom = GlobalAddAtom(strCommon)
            Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
            gobjCls(IntCount).CodeMan lngSys, Modul, gcnOracle, FrmMainface, gstrDbUser
            Call GlobalDeleteAtom(intAtom)
            
            '��ҽ������ֻ��CodeMan()���ܻ�ȡϵͳ�ţ��ڶ�ȡ����ʱ����֪��ϵͳ�ţ���д��ע������ҽ��������Ĭ��Ϊ 100
            Call SaveSetting("ZLSOFT", "����ȫ��", "ϵͳ��", lngSys)
        End If
    End If
    FrmMainface.MousePointer = 0
    Exit Sub
errH:
    FrmMainface.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub ReLogin()
    '����:���������¼
    mblnע�� = True
    
    Call gobjRelogin.ReLogin(FrmMainface)
End Sub

Public Function OwnerUser(ByVal strUserName As String) As Boolean
    Dim RecUser As New ADODB.Recordset
    Dim strSQL As String
    OwnerUser = True
    On Error GoTo errH
'        If .State = 1 Then .Close
        strSQL = "Select Count(*) ������ From ZlSystems Where ������='" & strUserName & "'"
         Set RecUser = zlDatabase.OpenSQLRecord(strSQL, "������")
'        .Open "Select Count(*) ������ From ZlSystems Where ������='" & strUserName & "'", gcnOracle By zq
        
        If RecUser.EOF Then
            If Not IsNull(RecUser!������) Then
                If RecUser!������ = 0 Then OwnerUser = False
            End If
        End If
'    End With
    Exit Function
errH:
    OwnerUser = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CreateSynonyms(ByVal lngSys As Long, ByVal lngModul As Long)
    Dim strSQL As String
    '����ģ����������ͬ���(����Ѵ����򲻻��ٴ���)
    On Error Resume Next
    If mlngSysPre <> lngSys Then
        strSQL = "Zl_Createsynonyms(" & lngSys & ")"
        zlDatabase.ExecuteProcedure strSQL, "����ͬ���"
        mlngSysPre = lngSys
    End If
End Function

Public Sub AddHistory(ByVal strModul As String)
    Dim strϵͳ As String, str��� As String, intMax As Integer
    Dim arrϵͳ As Variant, arr��� As Variant, strValue As String
    Dim intϵͳ_Cur As Integer, int���_Cur As Integer
    Dim intϵͳ_Max As Integer, int���_Max As Integer
    '������еĳ���ʼ���ڵ�һ��λ�ã�����Ѵ�������ʷ��¼�У��������ڵ�һ��λ��
    'strModul:ϵͳ & "," & ģ��
    
    intMax = 6
    
    strValue = zlDatabase.GetPara("���ʹ��ģ��")
    If UBound(Split(strValue, "|")) >= 1 Then
        strϵͳ = Trim(Split(strValue, "|")(0))
        str��� = Trim(Split(strValue, "|")(1))
    End If
    If strϵͳ = "" Or str��� = "" Then
        strϵͳ = Split(strModul, ",")(0)
        str��� = Split(strModul, ",")(1)
        Call zlDatabase.SetPara("���ʹ��ģ��", strϵͳ & "|" & str���)
        Exit Sub
    End If
    
    arrϵͳ = Split(strϵͳ, ",")
    arr��� = Split(str���, ",")
    intϵͳ_Max = UBound(arrϵͳ)
    int���_Max = UBound(arr���)
    strϵͳ = Split(strModul, ",")(0): str��� = Split(strModul, ",")(1)
    If intϵͳ_Max > intMax Then intϵͳ_Max = intMax
    
    For intϵͳ_Cur = 0 To intϵͳ_Max
        int���_Cur = intϵͳ_Cur
        If int���_Cur > int���_Max Then Exit For
        If Not (arrϵͳ(intϵͳ_Cur) = Split(strModul, ",")(0) And arr���(int���_Cur) = Split(strModul, ",")(1)) Then
            strϵͳ = strϵͳ & "," & arrϵͳ(intϵͳ_Cur)
            str��� = str��� & "," & arr���(int���_Cur)
        End If
    Next
    Call zlDatabase.SetPara("���ʹ��ģ��", strϵͳ & "|" & str���)
End Sub

Public Sub CheckWinVersion()
    Dim lngVersion As Long
    
    mblnע�� = False
    lngVersion = GetVersion()
    If ((lngVersion And &H80000000) = 0) Then
        mlngWin32 = WINDOWSNT
    Else
        mlngWin32 = WINDOWS95
    End If
End Sub

Public Sub ShutDown()
    If mblnע�� Then Exit Sub
    If Val(zlDatabase.GetPara("�ر�Windows")) = 0 Then Exit Sub
    If mlngWin32 = WINDOWSNT Then
        'ExitWindowsEx lng�رռ��������Դ Or EWX_FORCEIFHUNG, 0
        Call AdjustToken
        Call ExitWindowsEx(M_lng�رռ��������Դ Or EWX_FORCE, 0)
    Else
        Call ExitWindowsEx(M_lng�رռ��������Դ Or EWX_FORCE, 0)
    End If
End Sub

Public Function AdjustToken() As Boolean
    Const TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8
    Const SE_PRIVILEGE_ENABLED = &H2
    Dim hdlProcessHandle As Long
    Dim hdlTokenHandle As Long
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    
    'Set the error code of the last thread to zero using the'SetLast Error function
    SetLastError 0
    
    '�õ���ǰ���̵ľ��
    hdlProcessHandle = GetCurrentProcess()
    If GetLastError <> 0 Then Exit Function
    
    '�õ���ǰ���̵�Ȩ�޾��
    OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle
    If GetLastError <> 0 Then Exit Function
     
    '�ҵ��ر�Ȩ�޲�����LUID
    'SE_REMOTE_SHUTDOWN_NAME = "SeRemoteShutdownPrivilege
    'SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
    
    tkp.PrivilegeCount = 1    ' One privilege to set
    tkp.TheLuid = tmpLuid
    tkp.Attributes = SE_PRIVILEGE_ENABLED
    
    'Enable the shutdown privilege in the access token of this process
    AdjustTokenPrivileges hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
    If GetLastError <> 0 Then Exit Function
    
    AdjustToken = True
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim intDo As Integer
    Dim StrPass As String, strReturn As String, strSource As String, strTarget As String
    
    StrPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(StrPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function
