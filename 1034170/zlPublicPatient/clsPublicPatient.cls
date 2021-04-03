VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function zlInitCommon(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, _
      Optional ByVal strDBUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDBUser
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrUnitName = gobjComlib.GetUnitName
    Call InitLocPar
    
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ModiPatiBaseInfo(frmMain As Object, ByVal strģ�� As String, Optional ByVal lng����ID As Long = 0, _
    Optional ByVal lng����ID As Long = 0, Optional ByVal int���� As Integer = 1, Optional ByVal blnBatch As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:frmMain-����������
    '     strģ��-����ģ������,�磺����Һ�,����ҽ������վ
    '     lng����ID:����ID=0ͨ��-����ID�ȷ�ʽ�ֹ����˲��ˣ�����ID<>0ֱ����ȡ��Ӧ������Ϣ
    '     lng����ID=��0:�Һ�ID����ҳID(�����Զ���λ��Ҫ�޸ĵ�ĳһ��סԺ�����)��0��ʾ�û��ֹ�ѡ�������ﻹ��סԺ
    '     int���� 1-����;2-סԺ
    '     blnBatch:True:�����л����ˣ������������.False:���������ĳ�ξ���(�����л����˺�ѡ��������)
    '����:TRUE��False
    '����:������
    '����:2014-07-16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ModiPatiBaseInfo = frmModiPatiBaseInfo.ShowMe(frmMain, lng����ID, lng����ID, int����, strģ��, blnBatch)
End Function

Public Function PatiImageGatherer(frmMain As Object, ByRef strPictureFilePath As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ͼƬ�ɼ�
    '���:frmMain-����������
    '����:strPictureFilePath-���زɼ�ͼƬ��·��(�����ļ�������׺)
    '����:TRUE��False
    '����:������
    '����:2014-07-16
    '����ʾ��:�˽ӿڷ���TRue����ֱ��ʹ��LoadPicture(strPictureFilePath)����ͼƬ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    PatiImageGatherer = frmPatiImageGatherer.ShowMe(frmMain, strPictureFilePath)
End Function

Public Function CheckPatiAge(ByVal strAge As String, Optional ByVal strBirthday As String = "", _
    Optional ByVal bytTag As Byte = 0, Optional ByVal strCalcDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������
    '���:
    '   strAge:��������
    '   strBirthDay:��������
    '   bytTag:����zl_Age_Check�������ص�ѯ�����͵���Ϣ���Ƿ�Ҫǿ����ֹ�����Ǳ���ѯ��.0-����ѯ��,1-��ֹ
    '   strCalcDate:��������,ȱʡ���յ�ǰʱ�����
    '���أ�TRUE��FALSE��TRUE:����,FALSE:��ֹ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    CheckPatiAge = CheckAge(strAge, strBirthday, bytTag, strCalcDate)
End Function

Public Function ReCalcBirthDay(ByVal strAge As String, ByRef strDateOfBirth As String, Optional ByVal strCalcDate As String, Optional ByRef strMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ��������ȡ���˳�������
    '���:strAge:��������,�磺23�ꡢ1��2��
    'strCalcDate-�����������
    '����:����Ĳ��������ʽ��ȷ����㷵�س�������,���򷵻ؿ�
    '  strMsg-���ؾ�����Ϣ
    '��ȷ�����ʽ:X��[X��]��X��[X��]��X�졢XСʱ[X����]
    '    X��:X���ܴ���200,X��:X���ܴ���12,X��:X���ܴ���31,XСʱ:X���ܴ���24,X����:X���ܴ���59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    ReCalcBirthDay = RecalcBirth(strAge, strDateOfBirth, strCalcDate, strMsg)
    
End Function

Public Function GetPatiState(ByVal lng����ID As Long, ByRef byt״̬ As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡��ĳ��סԺ״̬(��Ժ��Ԥ��Ժ��)
    '���:lng����ID������Ψһ��ʶ
    '����:byt״̬������״̬:1-��Ժ��2-Ԥ��Ժ��3-������4-��Ժ
    '����:TRUE/FALSE
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = " Select b.��Ժ����, b.״̬, b.��Ժ��ʽ" & vbNewLine & _
        " From ������Ϣ a, ������ҳ b" & vbNewLine & _
        " Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.����id = [1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ���˵�ǰ״̬", lng����ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    If IsNull(rsTemp!��Ժ����) Then
        If Val(Nvl(rsTemp!״̬, 0)) = 3 Then
            byt״̬ = 2
        Else
            byt״̬ = 1
        End If
    Else
        If Nvl(rsTemp!��Ժ��ʽ) = "����" Then
            byt״̬ = 3
        Else
            byt״̬ = 4
        End If
    End If
    GetPatiState = True
End Function

Public Function CheckPatiExistBackLists(ByVal byt��Դ As Byte, ByVal lng����ID As Long, ByVal byt��ʾ��ʽ As Byte, ByRef bln�Ƿ���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡���Ƿ������ⲡ�ˣ�������ѯ�ʻ��ֹ��
    '���:
    '   byt��Դ��1-ԤԼ;2-�����ƺ����;3������(�˲���Ŀǰδʹ��ΪԤ��������������Ҫʹ��)
    '   lng����ID: ����Ψһ��ʶ
    '   byt��ʾ��ʽ����-��ʾ;2-��ֹ
    '����:
    '   bln�Ƿ������TRUE����,False��ֹ.
    '����:TRUE/FALSE
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strBackList As String
    Dim i As Integer
    
    strSQL = "Select ����ԭ�� From ���ⲡ�� Where ����ʱ�� is Null And ����ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "���ⲡ��", lng����ID)
    If rsTmp.EOF Then
        bln�Ƿ���� = True: CheckPatiExistBackLists = True: Exit Function
    End If
    
    With rsTmp
        i = 1
        strBackList = ""
        Do While Not .EOF
            If i > 5 Then Exit Do
            If i = 4 Then
                strBackList = strBackList & "    " & "����" & vbCrLf
            Else
                strBackList = strBackList & "    " & "�� " & Nvl(!����ԭ��) & vbCrLf
            End If
            i = i + 1
            .MoveNext
        Loop
    End With
    If byt��ʾ��ʽ = 1 Then
        If MsgBox("�˲��������ⲡ�������С�" & vbCrLf & vbCrLf & "ԭ��" & vbCrLf & vbCrLf & strBackList & vbCrLf & " �������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            bln�Ƿ���� = True
        Else
            bln�Ƿ���� = False
        End If
    Else
        Call MsgBox("�˲��������ⲡ�������С�" & vbCrLf & vbCrLf & "ԭ��" & vbCrLf & vbCrLf & strBackList & vbCrLf & " ���ܼ�����", vbInformation + vbOKOnly, gstrSysName)
        bln�Ƿ���� = False
    End If
    
    CheckPatiExistBackLists = True
End Function

Public Function GetPatiDiagnose(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal byt�������� As Byte, _
    Optional ByVal str������� As String = "", Optional ByVal str��¼��Դ As String = "", Optional ByRef str�����Ϣ As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���סԺ���������Ϣ
    '���:
    '   lng����ID: ����Ψһ��ʶ
    '   lng����ID: ��ҳID��ҽ��ID���Һ�ID
    '   byt�������ͣ�1-��ȡ������ϣ���-��ȡָ�����͵������Ϣ��������ȡָ��ҽ����Ӧ����ϣ�4-��ȡ���һ�����(����˳��Ϊ����Ժ����Ժ������)
    '   str������ͣ�������ʹ�,����������֮���ö��ŷָ�,��:2,3,12,13��(��Ϊ��)
    '   str��¼��Դ�������Դ��ʽ��Ĭ��Ϊ�ա������Դ��ʽ֮���ö��ŷָ��:3,2,1,4����¼��Դֵ��Ϊ1-4��
    ' ˵��:
    '   1��byt�������ͣ�������ҳID�ɴ���,2,4�����ݹҺ�ID�ɴ���������ҽ��ID��ѯֻ�ܴ�����
    '   2��str������ͣ�byt��������=1����str�������ֻ����1��11,Ϊ���൱���������="1,11"��byt��������=2��str������Ͳ���Ϊ�գ�byt��������=3��str�������Ĭ��Ϊ�գ�byt��������=4��str�������Ĭ��Ϊ�ա�
    '   3��str��¼��Դ��ֻ��һ����Դ����ʾ��ȡָ����Դ��ʽ����ϣ����������Դ����ʾ���ݴ������Դ��ʽ���ȼ���ȡ���(���ȼ�ѭ��Ϊ����˳��)��Ϊ�ձ�ʾ��Դ��ʽ=3,2,1,4
    '����:
    '   str�����Ϣ���������+�Ƿ����������ʾΪ(?)�����������֮���ö��ŷָ
    '����:TRUE/FALSE
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strReturn As String
    Dim arrCode, i As Integer
    
    If InStr(1, ",1,2,3,4,", "," & byt�������� & ",") = 0 Then Exit Function
    
    Select Case byt��������
        Case 3 '��ȡָ��ҽ����Ӧ�����
            strSQL = " Select a.�������, a.�Ƿ�����" & vbNewLine & _
                " From ������ϼ�¼ a, �������ҽ�� b" & vbNewLine & _
                " Where a.Id = b.���id And b.ҽ��id = [1] And ȡ��ʱ�� Is Null"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡָ��ҽ����Ӧ�����", lng����ID)
        Case Else
            strSQL = "Select ��¼��Դ, �������, ��ϴ���, �������, �Ƿ�����, Mod(�������, 10) As ����" & vbNewLine & _
                " From ������ϼ�¼" & vbNewLine & _
                " Where ����id = [1] And ��ҳid = [2] And NVL(�������,1) = 1 And Instr([3], ',' || ������� || ',', 1) <> 0" & vbNewLine & _
                " Order By ��¼��Դ, �������, ��ϴ���"
            If byt�������� = 1 Then '-�������
                If str������� = "" Then str������� = ",1,11,"
            ElseIf byt�������� = 2 Then '-��ȡָ�����Ͷ�Ӧ�����
                If str������� = "" Then
                    str�����Ϣ = "": Exit Function
                Else
                    str������� = "," & str������� & ","
                End If
            Else  '-��ȡ�������һ�����
                str�����Ϣ = ",1,2,3,11,12,13,"
            End If
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ�������", lng����ID, lng����ID, str�������)
            If str��¼��Դ = "" Then str��¼��Դ = "3,2,1,4"
            arrCode = Split(str��¼��Դ, ",")
            For i = 0 To UBound(arrCode)
                rsTmp.Filter = "��¼��Դ=" & Val(arrCode(i))
                If Not rsTmp.EOF Then Exit For
            Next i
            
            If byt�������� = 4 And Not rsTmp.EOF Then
                strSQL = rsTmp.Filter
                rsTmp.Filter = strSQL & " And ����=3"
                If rsTmp.EOF Then rsTmp.Filter = strSQL & " And ����=2"
                If rsTmp.EOF Then rsTmp.Filter = strSQL & " And ����=1"
            End If
    End Select
    
    strReturn = ""
    Do While Not rsTmp.EOF
        If Not IsNull(rsTmp!�������) Then
            strReturn = IIf(strReturn = "", "", strReturn & "��") & rsTmp!������� & IIf(Nvl(rsTmp!�Ƿ�����, 0) = 1, "������", "")
        End If
        rsTmp.MoveNext
    Loop
    
    str�����Ϣ = strReturn
    GetPatiDiagnose = True
End Function

Public Function HosPatiInfoCount(ByVal byt���� As Byte, ByVal lng����ID As Long, ByVal dtBegin As Date, ByVal dtEnd As Date, _
    ByVal str״̬�� As String, ByRef rsPatiCount As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:סԺ������Ϣͳ��
    '���:
    '   byt���ͣ�0-���ң�1-����
    '   lng����ID: ���һ���ID
    '   dtBegin: ͳ�ƵĿ�ʼʱ��?
    '   dtEnd: ͳ�ƵĽ���ʱ��?
    '   Str״̬����Ҫͳ����Щ״̬(�磺��Ժ��ת��)�Ĳ��ˣ���ʽΪ��0��1ƴ�ӳɵ��ַ���(0-��ͳ�ơ�1-ͳ��)���ָ���Ϊ�ֺ�";"��Ŀǰ֧������˳��Ϊ����Ժ����ơ�ת������Ժ��������Σ���ز��ˡ�˵�����ò����Ժ����չ��
    '����:
    '   rsPatiCount���ֶ����ƺ�˳������Ϊ����Ժ������ת�������ת����������Ժ������������������Σ����������������
    '����:TRUE/FALSE
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, rsPati As New ADODB.Recordset
    Dim strSQL As String
    With rsPati
        If .State = adStateOpen Then .Close
        .Fields.Append "��Ժ����", adDouble, 18, adFldIsNullable
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
        .Fields.Append "ת������", adDouble, 18, adFldIsNullable
        .Fields.Append "��Ժ����", adDouble, 18, adFldIsNullable
        .Fields.Append "��������", adDouble, 18, adFldIsNullable
        .Fields.Append "��Σ����", adDouble, 18, adFldIsNullable
        .Fields.Append "��������", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
'    '��Ժ����
'    strSQL = "" & _
'            " SELECT SUM(��Ժ) AS ��Ժ,SUM(���) AS ���,SUM(ת��) AS ת��,SUM(����) AS ����,SUM(��Ժ) AS ��Ժ,SUM(Σ) AS Σ,SUM(��) AS ��" & _
'            " FROM (" & _
'            "     SELECT SUM(DECODE(��ʼԭ��,2,1,0)) AS ��Ժ,SUM(DECODE(��ʼԭ��,3,1,15,1,0)) AS ���,0 AS ת��,0 AS ����,0 AS ��Ժ,0 AS Σ,0 AS ��" & _
'            "     From ���˱䶯��¼" & _
'            "     Where ����ID = [1] And NVL(���Ӵ�λ,0)=0" & _
'            "     AND ��ʼʱ�� BETWEEN [2] AND SYSDATE" & _
'            "     Union" & _
'            "     Select SUM(1) as ��Ժ,0 AS ���,0 AS ת��,0 AS ����,0 AS ��Ժ,0 AS Σ,0 AS ��" & _
'            "     From ���˱䶯��¼ a, ������ҳ b" & _
'            "     Where a.����id = b.����id And a.��ҳid = b.��ҳid And A.����ID=[1] And A.��ʼʱ�� Between [2] And Sysdate And a.��ʼԭ�� = 1 And Nvl(a.���Ӵ�λ, 0) = 0 And" & _
'            "       Nvl(b.״̬, 0) <> 1 And Not Exists" & _
'            "       (Select 1 From ���˱䶯��¼ Where ����id = a.����id And ��ҳid = b.��ҳid And ��ʼԭ�� = 2)"
'    strSQL = strSQL & _
'            "     Union" & _
'            "     SELECT 0 AS ��Ժ,0 AS ���,SUM(DECODE(��ֹԭ��,3,1,15,1,0)) AS ת��,0 AS ����,0 AS ��Ժ,0 AS Σ,0 AS ��" & _
'            "     From ���˱䶯��¼" & _
'            "     Where ����ID = [1] And NVL(���Ӵ�λ,0)=0" & _
'            "     AND ��ֹʱ�� BETWEEN [2] AND SYSDATE" & _
'            "     Union" & _
'            "     SELECT 0 AS ��Ժ,0 AS ���,0 AS ת��,SUM(DECODE(��Ժ��ʽ,'����',1,0)) AS ����,SUM(DECODE(��Ժ��ʽ,'����',0,1)) AS ��Ժ,0 AS Σ,0 AS ��" & _
'            "     From ������ҳ A,������Ϣ B" & _
'            "     Where A.����ID=B.����ID  And A.��ǰ����ID = [1]" & _
'            "     AND ��Ժ���� BETWEEN [2] AND SYSDATE" & _
'            "     Union" & _
'            "     SELECT 0 AS ��Ժ,0 AS ���,0 AS ת��,0 AS ����,0 AS ��Ժ,SUM(DECODE(��ǰ����,'Σ',1,0)) AS Σ,SUM(DECODE(��ǰ����,'��',1,0)) AS ��" & _
'            "     From ������ҳ A,������Ϣ B,��Ժ���� C" & _
'            "     Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And NVL(A.״̬,0)<>1 And Nvl(A.����״̬,0)<>5 And A.���ʱ�� is NULL And B.����ID=C.����ID " & _
'            "       And B.��ǰ����ID=C.����ID And C.����ID=[1]" & _
'            ")"
    HosPatiInfoCount = True
End Function

Public Function CheckPatiIdcard(ByVal strIdcard As String, Optional strBirthday As String, Optional strAge As String, Optional strSex As String, Optional strErrInfo As String, Optional datCalc As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����֤����Ϸ���У��
    '��Σ�strIdCard ����֤����
    '���Σ�strBirthday  ��������TrueΪ��������
    '         strSex ��������TrueΪ�Ա�
    '         strErrInfo ��������FalseΪ������Ϣ
    '         datCalc �������� ȱʡ��ϵͳʱ�����
    '���أ�True/False  ����֤�Ϸ�����True(�ɴ�strBirthday��strSex��ȡ�������ں��Ա�)�����򷵻�False(�ɴ�strErrInfo��ȡ��ϸ������Ϣ)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    CheckPatiIdcard = CheckIdcard(strIdcard, strBirthday, strAge, strSex, strErrInfo, datCalc)
End Function

Public Function SavePatiBaseInfo(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal strName As String, ByVal strSex As String, _
    ByVal strAge As String, ByVal strBirthday As String, ByVal strģ�� As String, Optional ByVal int���� As Integer = 1, Optional strInfo As String = "", _
    Optional ByVal blnXWHIS As Boolean = True, Optional ByVal blnEMPI As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ��������˻�����Ϣ(��ҵ�����ݵ�ͬ������)
    '��Σ�lng����ID-����ID (����Ϊ��/0)
    '         lng����ID-�Һ�ID����ҳID(��Ϊ0)
    '         strName-���� (����Ϊ��)
    '         strSex-�Ա� (����Ϊ��)
    '         strAge-���� (����Ϊ��)
    '         strBirthDay-�������� (����Ϊ��)
    '         strģ��-���øù��ܵ�ģ����������"����Һ�"��"��鱨��"��
    '         int���� 1-����;2-סԺ(lng����ID=0,��Ĭ��Ϊ1;lng����ID<>0,1-lng����IDΪ�Һ�ID,2-lng����IDΪ��ҳID)
    '         blnXWHIS-������Ϣ����ʱ�Ƿ����RIS�Ľӿ� ȱʡ=True���ã��ò������ڱ��ⲡ����Ϣ���ظ�����RIS�ӿڣ�
    '         blnEMPI-T EMPIƽ̨�Ѿ�������F-EMPIƽ̨δ����
    ' ���Σ�strInfo:���³ɹ�-��Ϣ�������µı仯��Ϣ(����True); ����ʧ��-��Ϣ����δ�ɹ���ԭ��
    ' ���أ�TRUE OR False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNote As String
    
    If lng����ID = 0 Then strInfo = "�������˻�����Ϣ������Ҫ��ȷ���ˡ�": Exit Function
    If Trim(strName) = "" Then strInfo = "�������˻�����Ϣ������Ҫ��ȷ����������": Exit Function
    If Trim(strSex) = "" Then strInfo = "�������˻�����Ϣ������Ҫ��ȷ�����Ա�": Exit Function
    If IsDate(strBirthday) = False Then strInfo = "�������˻�����Ϣ������Ҫ��ȷ���˳������ڡ�": Exit Function
    If Trim(strAge) = "" Then strInfo = "�������˻�����Ϣ������Ҫ��ȷ�������䡣": Exit Function
    
    Err.Clear
    On Error GoTo ErrHand
    SavePatiBaseInfo = SaveBaseInfo(lng����ID, lng����ID, strName, strSex, strAge, strBirthday, strģ��, int����, strInfo, blnXWHIS, blnEMPI)
    Exit Function
ErrHand:
    If gcnOracle.Errors.Count <> 0 Then
        strNote = gcnOracle.Errors(0).Description
        If InStr(UCase(strNote), "[ZLSOFT]") > 0 Then
            strInfo = Split(strNote, "[ZLSOFT]")(1)
        Else
            strInfo = strNote
        End If
    Else
        strInfo = Err.Description
    End If
End Function

Public Function ShowSelect(rsSelect As ADODB.Recordset, ByVal strKey As String, Optional ByVal strTitle As String, Optional ByVal strNote As String, _
    Optional ByVal blnMutilSelect As Boolean = False, Optional ByVal blnSerach As Boolean = False, Optional ByVal strMshWidth As String = "", Optional ByVal blnHideCancel As Boolean) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------------------
'����:������Ϣѡ����
'
'����       True=ȷ��,Flase=ȡ��
'-----------------------------------------------------------------------------------------------------------------------------------------------
   ShowSelect = frmListSel.ShowSelect(rsSelect, strKey, strTitle, strNote, blnMutilSelect, blnSerach, strMshWidth, blnHideCancel)
End Function

Public Function GetPatiByID(ByVal lngModel As Long, ByVal strIDNumber As String, Optional ByVal strPatiName As String, _
    Optional ByVal strPatiSex As String) As Long
'����:���ݲ�������֤��(����,�Ա�)��ȡ����id,����id�п����Ƕ��(BUG:131413)
'���:  lngModel-ģ���
'       strIDNumber-����֤��
'       strPatiName-��������
'       strPatiSex-�����Ա�


    Dim strSQL As String
    Dim rsTmp  As ADODB.Recordset
    Dim strPatiID As String
    Dim strMsg As String
    Dim lngPatiId As Long
    
    On Error GoTo errH
    strSQL = "Select zl_Custom_PatiIDs_Get([1],[2],[3],[4]) As ����IDs From dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPatiByID", lngModel, strIDNumber, strPatiName, strPatiSex)
    If rsTmp.EOF Then Exit Function
    strPatiID = Nvl(rsTmp!����IDs)
    If InStr(strPatiID, ",") > 0 Then

        strSQL = "Select Rownum As ID, ����id, ����, �Ա�, ����, ��������, ����֤��, �����, סԺ��, ��ַ, To_Char(�Ǽ�ʱ��, 'YYYY-MM-DD') As �Ǽ�����" & vbNewLine & _
            "From (Select /*+cardinality(B,10)*/" & vbNewLine & _
            "       Distinct a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'YYYY-MM-DD') As ��������, a.����֤��, a.�����, a.סԺ��," & vbNewLine & _
            "                Nvl(a.��ͥ��ַ, 'δ�Ǽ�') As ��ַ, a.����ʱ��, a.��Ժʱ��, a.�Ǽ�ʱ��" & vbNewLine & _
            "       From ������Ϣ A, Table(f_Str2list([1])) B" & vbNewLine & _
            "       Where a.����id = b.Column_Value " & vbNewLine & _
            "       Order By Nvl(Nvl(a.����ʱ��, a.��Ժʱ��), a.�Ǽ�ʱ��) Desc)"
            
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPatiByID", strPatiID)
        
        strMsg = " ��ǰ���������������Ϣ����ѡ��ָ���Ĳ���"
        If ShowSelect(rsTmp, "ID", "����ѡ��", strMsg, , , "0|1000|1000|500|800|1000|1800|1000|1000|1500|1200") Then
           lngPatiId = Val(rsTmp!����ID & "")
        Else
           lngPatiId = -1
        End If
    Else
        lngPatiId = Val(strPatiID)
    End If
    GetPatiByID = lngPatiId
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckPatiExistByID(ByVal strIDNumber As String, ByVal lngPatiId As Long) As Boolean
'����:���ݲ�������֤���ж��Ƿ������ͬ��������
'����ֵ:True -������ͬ����;False-��������ͬ����
    Dim strSQL As String
    Dim rsTmp  As ADODB.Recordset
    
    On Error GoTo errH

    strSQL = "Select ����id From ������Ϣ Where ����֤�� = [1] And ����id <> [2]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckPatiExistByID", strIDNumber, lngPatiId)
    CheckPatiExistByID = (rsTmp.RecordCount > 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckPatiIn(ByVal lngPatiId As Long) As Boolean
'����:���ݲ�������֤��(����,�Ա�)��ȡ����id,����id�п����Ƕ��(BUG:132712)
'���:  lngModel-ģ���
'       lngPatiID-����ID
'����ֵ:
'       True-���ھ�ҽ;False-���ھ�ҽ
    Dim strSQL As String
    Dim rsTmp  As ADODB.Recordset
    
    On Error GoTo errH
    '0-������Ժ��ҽ״̬��1-�Ǵ�����Ժ��ҽ״̬
    strSQL = "Select Zl_Pati_Is_Inhospital([1]) As �Ƿ��ҽ From dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckPatiIn", lngPatiId)
    If rsTmp.EOF Then Exit Function
    CheckPatiIn = (Val(rsTmp!�Ƿ��ҽ & "") = 1)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Class_Terminate()
    If Forms.Count = 0 Then
        Call CloseSquareCardObject '�رտ�������Ϣ
    End If
End Sub