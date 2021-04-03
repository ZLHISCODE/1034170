Attribute VB_Name = "mdlPublic"
Option Explicit
Public gintType As Integer
Public Enum ҽ��Enum
    TYPE_ͭ���� = 81
End Enum

'����˾ʵ�ֵ�DLL
Public Declare Function FTPUpLoad Lib "FTP_Trans.dll" (ByVal aHost As String, ByVal aPort As String, ByVal aUserID As String, ByVal aPassWord As String, ByVal aLocalFile As String, ByVal aRemoteDir As String, ByVal aRemoteFileName As String) As Long
Public Declare Function FTPDownLoad Lib "FTP_Trans.dll" (ByVal aHost As String, ByVal aPort As String, ByVal aUserID As String, ByVal aPassWord As String, ByVal aRemoteDir As String, ByVal aRemoteFileName As String, ByVal aLocalFile As String) As Long

Public Declare Function EncryptStr Lib "FTP_Trans.dll" (ByVal SourceStr As String, ByVal Key As String, ByVal IsEncrypt As Boolean) As String
Public Declare Function EncryptFiles Lib "FTP_Trans.dll" (ByVal INFName As String, ByVal OutFName As String) As Long
Public Declare Function DecryptFiles Lib "FTP_Trans.dll" (ByVal INFName As String, ByVal OutFName As String) As Long

Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public gcnOracle As New ADODB.Connection
Public gcnҽ�� As New ADODB.Connection
Public gstrSysName As String
Public gstrOwner As String
Public gstrSQL As String

Public Sub Main()
    Dim lngReturn As Long
    Dim strCode As String, IntCount As Integer, StrStyle As String
    Dim rsMenu As ADODB.Recordset, StrHaveSys As String
    
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��ʾ", "")
    If gstrSysName = "" Then gstrSysName = "�������"
    
    '�û�ע��
    frmUserLogin.Show 1
    If gcnOracle.State <> adStateOpen Then
        Unload frmUserLogin
        Exit Sub
    End If
    
    If ���ҽ�������� = False Then
        Exit Sub
    End If
    If ���ҽ�����ݱ� = False Then
        Exit Sub
    End If
    
    frm�ϴ�����.Show
    
    
End Sub

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
    Dim rsTemp As New ADODB.Recordset

    On Error Resume Next
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
        
        gstrOwner = UCase(strUserName)
        gstrSQL = "Select ��� From zlsystems where ������='" & gstrOwner & "' and trunc(���/100) in (1,8)"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        If rsTemp.RecordCount = 0 Then
            MsgBox "��¼�û�������ϵͳ�����ߡ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        .Execute "select * from ���˷��ü�¼ where rownum<1"
        If Err <> 0 Then
            MsgBox "�㲻���з���HIS���ݱ��Ȩ�ޡ�", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    OraDataOpen = True
End Function

Private Function ���ҽ��������() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    '��������ҽ��������������
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=" & TYPE_ͭ����
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "ҽ���û���"
                strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ��������"
                strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ���û�����"
                strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                '����
                If strPass <> "" Then strPass = EncryptStr(strPass, 256, False)
        End Select
        rsTemp.MoveNext
    Loop
    
    On Error Resume Next
    gcnҽ��.Provider = "MSDataShape"
    gcnҽ��.Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer, strUser, strPass
    
    If Err <> 0 Then
        MsgBox "ҽ��ǰ�÷���������ʧ�ܡ�" & vbCrLf & _
               "��ע�⣬���������������еķ������������л�����Ӧ����ͬ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    ���ҽ�������� = True
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

Public Function Currentdate() As Date
'���ܣ���õ�ǰ����
    Dim rsTmp As New ADODB.Recordset
    On Error Resume Next
    rsTmp.Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    Currentdate = rsTmp.Fields(0).Value
    If Err <> 0 Then
        '�õ�ǰ����ʱ��
        Currentdate = date
        Err.Clear
    End If
End Function

Public Function MidUni(ByVal strTemp As String, ByVal Start As Long, ByVal Length As Long) As String
'���ܣ������ݿ����õ��ַ������Ӽ���Ҳ���Ǻ��ְ������ַ��㣬����ĸ����һ��
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
End Function


Public Function AddDate(ByVal strOrin As String) As String
'���ܣ�Ϊ��ȫ��������Ϣ��������
    Dim strTemp As String
    Dim intPos As Integer
    
    strTemp = Trim(strOrin)
    
    If strTemp = "" Then
        AddDate = ""
        Exit Function
    End If
    
    intPos = InStr(strTemp, "-")
    If intPos = 0 Then
        intPos = InStr(strTemp, ".")
        If intPos <> 0 Then
            'ʹ�� . ��
            strTemp = Replace(strTemp, ".", "-")
        End If
    End If
    
    If intPos = 0 Then
        'û��"-",�ֹ�����
        intPos = Len(strTemp)
        If intPos <= 8 Then
            If intPos = 8 Then
                strTemp = Mid(strTemp, 1, 4) & "-" & Mid(strTemp, 5, 2) & "-" & Mid(strTemp, 7, 2)
            ElseIf intPos > 4 Then
                strTemp = Left(strTemp, intPos - 4) & "-" & Mid(Right(strTemp, 4), 1, 2) & "-" & Right(strTemp, 2)
            ElseIf intPos > 2 Then
                strTemp = Format(date, "yyyy") & "-" & Left(strTemp, intPos - 2) & "-" & Right(strTemp, 2)
            Else
                strTemp = Format(date, "yyyy") & "-" & Format(date, "MM") & "-" & strTemp
            End If
        End If
    Else
        If IsDate(strTemp) Then
            strTemp = Format(CDate(strTemp), "yyyy-MM-dd")
        End If
    End If
    
    AddDate = strTemp
End Function

Private Function ���ҽ�����ݱ�() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    'ֻ����ϴ����ر�
    On Error Resume Next
    gstrSQL = "select * from �ϴ����� where rownum<1"
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    
    If Err <> 0 Then
        MsgBox "������������õ�ҽ���û������߱���ҽ���йص����ݱ������а�װ�ű���", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�������
    On Error GoTo errHandle
    
    
    
    ���ҽ�����ݱ� = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0, Optional ByVal hwnd As Long = 0, Optional str��Ŀ As String) As Boolean
'����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If str��Ŀ = "" Then str��Ŀ = "����������"
    
    If InStr(strInput, "'") > 0 Or InStr(strInput, "|") > 0 Then
        MsgBox str��Ŀ & "���зǷ��ַ���", vbExclamation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox str��Ŀ & "���ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, gstrSysName
            If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
            Exit Function
        End If
    End If
    
    StrIsValid = True
End Function


Public Sub OpenRecordset(rsTemp As ADODB.Recordset, _
        Optional CursorType As CursorTypeEnum = adOpenStatic, Optional LockType As LockTypeEnum = adLockReadOnly)
'���ܣ��򿪼�¼��ͬʱ����SQL���
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    rsTemp.Open gstrSQL, gcnҽ��, CursorType, LockType
End Sub

Public Sub ExecuteProcedure()
'���ܣ�ִ�й���ʽ��SQL���
    gcnҽ��.Execute gstrSQL, , adCmdStoredProc
End Sub

Public Function GetMax(ByVal strTable As String, ByVal strField As String, ByVal intLength As Integer, Optional ByVal strWhere As String) As String
'���ܣ���ȡָ����ı�����������ֵ
'������strTable  ����;
'      strField  �ֶ���;
'      intLength �ֶγ���
'���أ��ɹ����� �¼�������; ���߷��� 0
    Dim rsTemp As New ADODB.Recordset
    Dim varTemp As Variant, strSQL As String
    Dim lngLengh As Long
    
    On Error GoTo ErrHand
    With rsTemp
        strSQL = "SELECT MAX(LPAD(" & strField & "," & intLength & ",' ')) as ""���ֵ"",max(length(" & _
             strField & ")) as ""�ֵ"" FROM " & strTable & strWhere
        rsTemp.Open strSQL, gcnҽ��, adOpenStatic, adLockReadOnly
        
        If rsTemp.EOF Then
            GetMax = Format(1, String(intLength, "0"))
            Exit Function
        End If
        varTemp = IIf(IsNull(.Fields("���ֵ").Value), "0", .Fields("���ֵ").Value)
        lngLengh = IIf(IsNull(.Fields("�ֵ").Value), intLength, .Fields("�ֵ").Value)
        If IsNumeric(varTemp) Then
            GetMax = CStr(Val(varTemp) + 1)
            GetMax = Format(GetMax, String(lngLengh, "0"))
        Else
            GetMax = Mid(varTemp, 1, Len(varTemp) - 1) & Chr(Asc(Right(varTemp, 1)) + 1)
            GetMax = Trim(GetMax)
        End If
        .Close
    End With
    Exit Function
    
ErrHand:
    If frmErr.ShowErr(Err.Description) = vbYes Then Resume
End Function


