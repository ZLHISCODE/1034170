VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCISMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mint���� As Integer '0-����ҽ������վ��1��סԺҽ������վ��2��סԺ��ʿ����վ��3���ϰ�ҽ������վ
Private mstrCurDate As String '��ǰʱ��
 
Private mobjVBA As Object
Private mobjScript As clsScript
Private mobjVoice As Object                  '�������ж���
Private mrsPars As ADODB.Recordset '������Ϣ
Private mrsMsgNo As ADODB.Recordset '��Ϣ���ͺͱ�Ŷ�Ӧ

Public Sub InitCISMsg(ByVal intType As Integer)
'���ܣ���ʼ��������
    Dim strTmp As String
    Dim varTmp As Variant
    
    Dim strNos As String
    Dim varNo As Variant
    
    Dim i As Long, lngģ�� As Long
    
    mint���� = intType
    
    Call CreateScript(mobjVBA, mobjScript)
    
    Call InitRsMsgPar(mrsPars)
    
    Set mrsMsgNo = New ADODB.Recordset
    mrsMsgNo.Fields.Append "��������", adVarChar, 800
    mrsMsgNo.Fields.Append "��Ϣ���", adVarChar, 60
    mrsMsgNo.CursorLocation = adUseClient
    mrsMsgNo.LockType = adLockOptimistic
    mrsMsgNo.CursorType = adOpenStatic
    mrsMsgNo.Open
    
    If mint���� = 0 Then
        lngģ�� = p����ҽ��վ
    ElseIf mint���� = 1 Then
        lngģ�� = pסԺҽ��վ
    ElseIf mint���� = 2 Then
        lngģ�� = pסԺ��ʿվ
    ElseIf mint���� = 3 Then
        lngģ�� = pҽ������վ
    End If
     
    strTmp = Get��Ϣ���(intType)
    varTmp = Split(strTmp, ",")
    strNos = Get��Ϣ���(intType)
    varNo = Split(strNos, "|")
    For i = 0 To UBound(varTmp)
        Call AddDataToRsMsgPar(mrsPars, lngģ��, i + 1, varTmp(i) & "��������", varTmp(i))
        Call AppendNoPar(varTmp(i), varNo(i))
    Next
    
    mstrCurDate = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
End Sub
 
Private Sub AppendNoPar(ByVal str��Ϣ�� As String, ByVal strNos As String)
'���ܣ���Ϣ����Ϣ��Ź���
    Dim varTmp As Variant
    Dim i As Long
    varTmp = Split(strNos, ",")
    For i = 0 To UBound(varTmp)
        mrsMsgNo.AddNew
        mrsMsgNo!�������� = str��Ϣ��
        mrsMsgNo!��Ϣ��� = varTmp(i)
        mrsMsgNo.Update
    Next
End Sub

Public Sub PlayMsgSound(ByRef rsMsg As ADODB.Recordset)
'������Ϣ
    Dim strBegin As String, strEnd As String
    Dim rs���� As ADODB.Recordset
    Dim i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    strBegin = mstrCurDate
    mstrCurDate = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    strEnd = mstrCurDate
    
    If rsMsg.RecordCount > 0 Then
        Set rs���� = zlDatabase.CopyNewRec(rsMsg, , , Array("�Ƿ񲥷�", adInteger, 10, 0))
        For i = 1 To rs����.RecordCount
            strTmp = Format(rs����!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss")
            If Between(strTmp, strBegin, strEnd) Then
                rs����!�Ƿ񲥷� = 1
            End If
            rs����.MoveNext
        Next
        rs����.Filter = "�Ƿ񲥷� = 1"
        If Not rs����.EOF Then
            For i = 1 To rs����.RecordCount
                Call PlaySound(rs����)
                rs����.MoveNext
            Next
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PlaySound(ByRef rsMsg As ADODB.Recordset)
'���ܣ���ʿվ��ָ�����ò���Ϣ
    Dim strText As String, strFile As String
    Dim int��ʾ��ʽ As Integer, i As Long
    Dim start As Date
    Dim str���� As String
    
    On Error GoTo errH
    
    If mobjVoice Is Nothing Then
        Set mobjVoice = CreateObject("SAPI.SpVoice")
    End If
   
    mrsMsgNo.Filter = "��Ϣ���='" & rsMsg!���ͱ��� & "'"
    If mrsMsgNo.EOF Then Exit Sub
    mrsPars.Filter = "��������='" & mrsMsgNo!�������� & "'"
    
    If Val(mrsPars!״̬ & "") = 0 Then Exit Sub
    
    str���� = mrsPars!���� & ""
    int��ʾ��ʽ = Val(mrsPars!��ʾ��ʽ & "")
    If int��ʾ��ʽ = 0 Then
        strText = Get�����ı�(rsMsg, str����)  '�ı�
    Else
        strFile = str���� '�ļ�
    End If
     
    For i = 1 To Val(mrsPars!���� & "")
        If int��ʾ��ʽ = 0 Then
            mobjVoice.Speak strText, 1
        Else
            Call sndPlaySound(strFile, 1) '�ļ���ʧ�𻵲��ᱨ������������
        End If
        start = Timer
        Do While Timer < start + 5
            Call Sleep(5)
            DoEvents
        Loop
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get�����ı�(ByRef rsMsg As ADODB.Recordset, ByVal strText As String) As String
'���ܣ���ȡ������ı�
    Dim str���� As String, strסԺ�� As String
    Dim strVal As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim str��Ŀ���� As String

    On Error GoTo errH
    str���� = IIf("" = rsMsg!���� & "", "", rsMsg!���� & "��")
    strסԺ�� = IIf("" = rsMsg!סԺ�� & "", "", rsMsg!סԺ�� & "��")
    strVal = strText
    strVal = Replace(strVal, "[����]", """" & str���� & """")
    strVal = Replace(strVal, "[סԺ��]", """" & strסԺ�� & """")
     
    If InStr("ZLHIS_PACS_006,ZLHIS_PACS_007", rsMsg!���ͱ��� & "") > 0 Then
        strSQL = "select ҽ������ as ���� from ����ҽ����¼ where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��Ϣ����", Val(rsMsg!ҵ���ʶ & ""))
        If Not rsTmp.EOF Then str��Ŀ���� = rsTmp!���� & ""
    End If
    strVal = Replace(strVal, "[��Ŀ����]", """" & str��Ŀ���� & """")
    
    strVal = mobjVBA.Eval(strVal)
    If mobjVBA.Error.Number <> 0 Then
        err.Clear
        strVal = IIf(str���� = "", "��ͥ��", str����) & rsMsg!��Ϣ����
    End If
    Get�����ı� = strVal
    Exit Function
errH:
    err.Clear
    Get�����ı� = "������  " & rsMsg!��Ϣ����
End Function

Private Sub Class_Terminate()
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    Set mobjVoice = Nothing
    Set mrsPars = Nothing
    Set mrsMsgNo = Nothing
End Sub

'-------------------���÷���
Public Sub InitRsMsgPar(ByRef rsPar As ADODB.Recordset)
'���ܣ���ʼ��������¼��
    Set rsPar = New ADODB.Recordset
    
    rsPar.Fields.Append "���", adBigInt
    rsPar.Fields.Append "ģ��", adBigInt
    rsPar.Fields.Append "������", adVarChar, 1000
    rsPar.Fields.Append "��������", adVarChar, 800
    rsPar.Fields.Append "״̬", adInteger '1-������0���ر�
    rsPar.Fields.Append "��ʾ��ʽ", adInteger '1����Ƶ�ļ���0���ʶ��ı�
    rsPar.Fields.Append "����", adVarChar, 3800
    rsPar.Fields.Append "����", adInteger
    rsPar.Fields.Append "�޸�", adInteger '1-���޸ģ�0��δ�޸�
    
    rsPar.Fields.Append "ԭ����ֵ", adVarChar, 4000
    rsPar.Fields.Append "�ֲ���ֵ", adVarChar, 4000
    
    rsPar.CursorLocation = adUseClient
    rsPar.LockType = adLockOptimistic
    rsPar.CursorType = adOpenStatic
    rsPar.Open
End Sub

Public Sub AddDataToRsMsgPar(ByRef rsPar As ADODB.Recordset, ByVal lngModual As Long, ByVal lng��� As Long, ByVal str������ As String, ByVal str��Ϣ�� As String)
'���ܣ��򱾵ؼ�¼�������Ӳ�������
    Dim varTmp As Variant
    Dim strTmp As String
    
    On Error GoTo errH
    strTmp = zlDatabase.GetPara(str������, glngSys, lngModual, "1<sTab>0<sTab>iif([����]<>"""",[����],""��ͥ��"")+""��" & str��Ϣ�� & "��""<sTab>2")
    varTmp = Split(strTmp, "<sTab>")
    rsPar.AddNew
    rsPar!��� = lng���
    rsPar!ģ�� = lngModual
    rsPar!������ = str������
    rsPar!�������� = str��Ϣ��
    rsPar!״̬ = Val(varTmp(0))
    rsPar!��ʾ��ʽ = Val(varTmp(1))
    rsPar!���� = varTmp(2)
    rsPar!���� = Val(varTmp(3))
    rsPar!ԭ����ֵ = strTmp
    rsPar!�ֲ���ֵ = strTmp
    rsPar!�޸� = 0
    rsPar.Update
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CreateScript(Optional ByRef objVBA As Object, Optional ByRef objScript As clsScript) As Boolean
'���ܣ�����Script��VBA����
    On Error Resume Next
    Set objVBA = CreateObject("ScriptControl")
    err.Clear: On Error GoTo 0
    If Not objVBA Is Nothing Then
        objVBA.Language = "VBScript"
        Set objScript = New clsScript
        objVBA.AddObject "clsScript", objScript, True
        CreateScript = True
    End If
End Function

Public Function Get��Ϣ���(ByVal intType As Integer) As String
'���ܣ���ȡÿ����Ϣ��Ӧ�ı��
'������intType '0-����ҽ������վ��1��סԺҽ������վ��2��סԺ��ʿ����վ��3���ϰ�ҽ������վ
    If intType = 0 Then
        Get��Ϣ��� = "ZLHIS_LIS_003,ZLHIS_PACS_005|ZLHIS_CIS_032,ZLHIS_CIS_033|ZLHIS_RECIPEAUDIT_001"
    ElseIf intType = 1 Then
        Get��Ϣ��� = "ZLHIS_EMR_021|ZLHIS_OPER_001,ZLHIS_CIS_005,ZLHIS_CIS_015|ZLHIS_LIS_003,ZLHIS_PACS_005|ZLHIS_LIS_002,ZLHIS_PACS_003|ZLHIS_CIS_026,ZLHIS_CIS_027,ZLHIS_CIS_028,ZLHIS_CIS_029,ZLHIS_CIS_030|ZLHIS_RECIPEAUDIT_002|ZLHIS_CIS_032,ZLHIS_CIS_033|ZLHIS_EMR_025"
    ElseIf intType = 2 Then
        Get��Ϣ��� = "ZLHIS_CIS_001|ZLHIS_CIS_002|ZLHIS_CIS_003|ZLHIS_OPER_001,ZLHIS_CIS_005,ZLHIS_CIS_015|ZLHIS_LIS_003,ZLHIS_PACS_005|ZLHIS_TRANSFUSION_001|ZLHIS_CHARGE_001|ZLHIS_PACS_006|ZLHIS_PACS_007"
    ElseIf intType = 3 Then
        Get��Ϣ��� = "ZLHIS_CHARGE_001|ZLHIS_CIS_004"
    End If
End Function

Public Function Get��Ϣ���(ByVal intType As Integer) As String
'���ܣ���ȡ����վ����Ϣ���
'������intType  0-����ҽ������վ��1��סԺҽ������վ��2��סԺ��ʿ����վ��3���ϰ�ҽ������վ
    If intType = 0 Then
        Get��Ϣ��� = "Σ��ֵ��Ϣ,��Ⱦ����Ϣ,���������Ϣ"
    ElseIf intType = 1 Then
        Get��Ϣ��� = "����������Ϣ,������Ϣ,Σ��ֵ��Ϣ,���泷����Ϣ,ҽ�������Ϣ,���������Ϣ,��Ⱦ����Ϣ,�����ʿ���Ϣ"
    ElseIf intType = 2 Then
        Get��Ϣ��� = "�¿���Ϣ,��ͣ��Ϣ,�·���Ϣ,������Ϣ,Σ��ֵ��Ϣ,��Һ�ܾ���Ϣ,����������Ϣ,RISԤԼ��Ϣ,RISԤԼ׼����Ϣ"
    ElseIf intType = 3 Then
        Get��Ϣ��� = "����������Ϣ,��������Ϣ"
    End If
End Function