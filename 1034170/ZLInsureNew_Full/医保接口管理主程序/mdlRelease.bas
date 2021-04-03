Attribute VB_Name = "mdlRelease"
Option Explicit
Public Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Public Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Public Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����
Public Enum ����
    �����֤
    �����֤_�����Һ�
    �ʻ����
    ����Һ�
    ����Һ�����
    �����������
    �������
    �����������
    �����ʻ�תԤ��
    Ԥ���˸����ʻ�
    סԺ�������
    סԺ����
    סԺ��������
    ��Ժ�Ǽ�
    ��Ժ�Ǽǳ���
    ��Ժ�Ǽ�
    ��Ժ�Ǽǳ���
    ������ϸ�ϴ�
    סԺ��Ϣ�䶯
    ��ȡҽ����Ŀ��Ϣ
    ��ȡҽ����Ŀ�����Ϣ
    ����ѡ��
    ȡ������Ǽ�
    ����
End Enum

Sub Main()
    frmUserLogin.Show 1
    If gcnOracle.State = 0 Then Exit Sub
    
    Call InitCommon(gcnOracle)
    frmҽ����������.Show
End Sub

Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Public Function CopyNewRec(ByVal SourceRec As ADODB.Recordset) As ADODB.Recordset
    Dim RecTarget As New ADODB.Recordset
    Dim intFields As Integer, LngLocate As Long
    '������:����
    '��������:2000-11-02
    '�ü�¼����ƾ֤�ؼ���Ӧ
    'Ҳʹ���ڱ���
    
    LngLocate = -1
    Set RecTarget = New ADODB.Recordset
    With RecTarget
        If .State = 1 Then .Close
        If SourceRec.RecordCount <> 0 Then
            On Error Resume Next
            Err = 0
            LngLocate = SourceRec.AbsolutePosition
            If Err <> 0 Then LngLocate = -1
            SourceRec.MoveFirst
        End If
        For intFields = 0 To SourceRec.Fields.Count - 1
            .Fields.Append SourceRec.Fields(intFields).Name, SourceRec.Fields(intFields).Type, SourceRec.Fields(intFields).DefinedSize, adFldIsNullable     '0:��ʾ����
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        If SourceRec.RecordCount <> 0 Then SourceRec.MoveFirst
        Do While Not SourceRec.EOF
            .AddNew
            For intFields = 0 To SourceRec.Fields.Count - 1
                .Fields(intFields) = SourceRec.Fields(intFields).Value
            Next
            .Update
            SourceRec.MoveNext
        Loop
    End With
    
    If SourceRec.RecordCount <> 0 Then SourceRec.MoveFirst
    If LngLocate > 0 Then SourceRec.Move LngLocate - 1
    Set CopyNewRec = RecTarget
End Function
