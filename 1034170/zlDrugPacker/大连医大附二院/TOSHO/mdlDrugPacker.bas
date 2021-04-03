Attribute VB_Name = "mdlDrugPacker"
Option Explicit

Public Const GSTR_MESSAGE = "��ʾ��Ϣ"

Public gstrUser As String, gstrUserNameNew As String
Public gintUserID As Integer, gintDeptID As Integer
Public gbytЧ�� As Byte

Public gobjComLib As Object                         'zl9Comlib����
Public gcnOutside As New ADODB.Connection           '�ⲿ���ݿ�����

Public Const GSTR_SYSNAME = "�Զ��ְ����ӿ�"
Public Const GSTR_REGEDIT_PATH = "����ģ��\DrugPackerDBServer"
Public Const MSTR_SERVER = "localhost"
Public Const MSTR_DBNAME = "atf"
Public Const MSTR_USER = "sa"
Public Const MSTR_PASSWORD = ""


Public Function MSSQLServerOpen(ByVal strServerName As String, ByVal strDBName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ����MS SQL Server ���ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    
    If Len(Trim(strUserName)) = 0 Then
        MSSQLServerOpen = False
        MsgBox "�������������ݿ���Ϣ��", vbInformation, GSTR_MESSAGE
        Exit Function
    End If
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOutside
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .ConnectionTimeout = 5
        .Open "Driver={SQL Server};Server=" & strServerName & ";Database=" & strDBName, strUserName, strUserPwd
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, GSTR_SYSNAME
            ElseIf Err.Number = -2147217843 Or Err.Number = -2147467259 Then
                MsgBox "ҩƷ�ְ������ݿ�����ʧ�ܣ�", vbInformation, GSTR_SYSNAME
            Else
                MsgBox strError, vbInformation, GSTR_SYSNAME
            End If
            
            MSSQLServerOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    'gstrDbUser = UCase(strUserName)
    'SetDbUser gstrDbUser
    
    MSSQLServerOpen = True
    Exit Function
    
errHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    MSSQLServerOpen = False
    Err = 0
End Function


Public Function OraDataOpenTest(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
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
    With gcnOutside
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, GSTR_SYSNAME
            ElseIf Err.Number = -2147217843 Then
                MsgBox Mid(strError, InStr(1, strError, "[SQL Server]"), Len(strError)), vbInformation, GSTR_SYSNAME
            Else
                MsgBox strError, vbInformation, GSTR_SYSNAME
            End If
            
            OraDataOpenTest = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    'gstrDbUser = UCase(strUserName)
    'SetDbUser gstrDbUser
    
    OraDataOpenTest = True
    Exit Function
    
errHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    OraDataOpenTest = False
    Err = 0
End Function

Public Function StringEnDeCodecn(strSource As String, MA) As String
'�ú���ֻ���������𵽼�������
'����Ϊ��Դ�ļ�������
    On Error GoTo ErrEnDeCode
    Dim X As Single, i As Integer
    Dim CHARNUM As Long, RANDOMINTEGER As Integer
    Dim SINGLECHAR As String * 1
    Dim strTmp As String
    
    If MA < 0 Then
        MA = MA * (-1)
    End If
    
    X = Rnd(-MA)
    For i = 1 To Len(strSource) Step 1                 'ȡ���ֽ�����
        SINGLECHAR = Mid(strSource, i, 1)
        CHARNUM = Asc(SINGLECHAR)
g:
        RANDOMINTEGER = Int(127 * Rnd)
        If RANDOMINTEGER < 30 Or RANDOMINTEGER > 100 Then GoTo g
        CHARNUM = CHARNUM Xor RANDOMINTEGER
        strTmp = strTmp & Chr(CHARNUM)
    Next i
    StringEnDeCodecn = strTmp
    Exit Function

ErrEnDeCode:
    StringEnDeCodecn = ""
    MsgBox Err.Number & "\" & Err.Description
End Function

Public Function GetUserNameInfo() As Boolean
'��ȡ�û���Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Set rsTmp = gobjComLib.GetUserInfo
    
    With rsTmp
        If Not .EOF Then
            gintUserID = IIf(IsNull(!Id), 0, !Id)
            gintDeptID = IIf(IsNull(!����id), 0, !����id)
            gstrUserNameNew = IIf(IsNull(!����), "", !����) '��ǰ�û�����
            GetUserNameInfo = True
        Else
            gintUserID = 0
            gintDeptID = 0
            gstrUserNameNew = "" '��ǰ�û�����
        End If
    End With
    rsTmp.Close

    strSQL = "Select ������, ����ֵ, ȱʡֵ From Zlparameters Where ϵͳ = [1] And Nvl(˽��, 0) = 0 And ģ�� Is Null and ������=[2] "
    Set rsTmp = gobjComLib.OpenSQLRecord(strSQL, "ȡϵͳ����", 100, 149)
    With rsTmp
        If Not .EOF Then
            gbytЧ�� = IIf(IsNull(rsTmp!����ֵ), rsTmp!ȱʡֵ, rsTmp!����ֵ)
        Else
            gbytЧ�� = 0
        End If
    End With
    
End Function
'
'Public Function CheckProvider(ByVal intProvider As Integer) As String
''��˹�Ӧ��ID
'    Dim rsTmp As New ADODB.Recordset
'    Set rsTmp = zlDatabase.OpenSQLRecord("select ���� from ��Ӧ�� where id=[1]", "��˹�Ӧ��ID", intProvider)
'    If rsTmp.RecordCount = 1 Then
'        CheckProvider = rsTmp!����
'    End If
'    rsTmp.Close
'End Function

Public Sub SelText(ByVal ctlVal As Control)
    If TypeOf ctlVal Is TextBox Then
        ctlVal.SelStart = 0
        ctlVal.SelLength = Len(ctlVal.Text)
    End If
End Sub



