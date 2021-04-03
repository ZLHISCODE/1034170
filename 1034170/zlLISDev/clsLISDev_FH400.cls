VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_FH400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ȫ�Զ�������
'�ͺ�:FH-400
'����:�Ϻ����
'�û�:�ձ�������ҽԺ
'ʱ�䣺2013/12/19
Private mstrCode As String

Public Function GetStartCmd() As String
'�豸��׼�ӿ�
'���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = Chr(6)
End Function

Public Function GetEndCmd() As String
'�豸��׼�ӿ�
'���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Function GetAnswerCmd() As String
'��ʱ���͵�ָ��
    GetAnswerCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
                   strCmd As String)
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    Dim str���� As String
    Dim str��� As String, str�걾�� As String
    Dim strͨ���� As String, lng_S As Long
    Dim lng_E As Long, iRow As Long, str_Code As String
    Dim varLine As Variant, strLine As String
    Dim aResult As String, varline2 As Variant
    Dim strline2 As String, j As Long
    Dim str_�����˱걾 As String

    str_�����˱걾 = " |Ѫ��|0"
    On Error GoTo errH
    
    If strSource <> "" Then
        strCmd = Chr(&H6&)
    End If
    mstrCode = mstrCode & strSource
    
    If InStr(Mid(mstrCode, 1, 1), Chr(&H4E)) > 0 Then     '�ж���ͨ�걾��H4E��ͷ,������������Ŀ�ʼλ��
        lng_S = InStr(mstrCode, Chr(&H4E))
    ElseIf InStr(Mid(mstrCode, 1, 1), Chr(&H45)) > 0 Then    '�жϼ���걾��H45��ͷ,������������Ŀ�ʼλ��
        lng_S = InStr(mstrCode, Chr(&H45))
    ElseIf InStr(Mid(mstrCode, 1, 1), Chr(&H43)) > 0 Then    '�ж��ʿر걾��H43��ͷ,������������Ŀ�ʼλ��
        lng_S = InStr(mstrCode, Chr(&H43))
    End If
    
    If lng_S > 0 Then
        lng_E = InStr(mstrCode, Chr(&HD) & Chr(&HA))      '�ж���HD����HA����,������������Ľ���λ��
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
        varLine = Split(str_Code, Chr(&HD))  '��ȡÿ������
        For iRow = LBound(varLine) To UBound(varLine)
            strLine = varLine(iRow)
            
            If InStr(strLine, Chr(&H4E)) > 0 Then
                varline2 = Split(str_Code, Chr(&H9))(0)  'ÿ��������H9����
                str���� = Format(Now, "yyyy-mm-dd hh:mm:ss")
                str�걾�� = Mid(Split(str_Code, Chr(&H9))(0), 2)    'Replace(Mid(strLine, 2, 3), Chr(&H9), "")
                strͨ���� = Split(str_Code, Chr(&H9))(1)
                str��� = Split(str_Code, Chr(&H9))(3)
            ElseIf InStr(strLine, Chr(&H45)) > 0 Then
                varline2 = Split(str_Code, Chr(&H9))(0)
                str���� = Format(Now, "yyyy-mm-dd hh:mm:ss")
                str�걾�� = Mid(Split(str_Code, Chr(&H9))(0), 2) + 1000
                strͨ���� = Split(str_Code, Chr(&H9))(1)
                str��� = Split(str_Code, Chr(&H9))(3)
            ElseIf InStr(strLine, Chr(&H43)) > 0 Then
                varline2 = Split(str_Code, Chr(&H9))(0)
                str���� = Format(Now, "yyyy-mm-dd hh:mm:ss")
                str�걾�� = Mid(Split(str_Code, Chr(&H9))(0), 2)
                strͨ���� = Split(str_Code, Chr(&H9))(1)
                str��� = Split(str_Code, Chr(&H9))(3)
            End If
            
        Next
        
        If str��� <> "" And str�걾�� <> "" Then
            '���ƴ���
            aResult = aResult & "||" & str���� & "|" & str�걾�� & "|" & str_�����˱걾 & "|" & strͨ���� & "|" & str���
        End If
        
        If InStr(Mid(mstrCode, 1, 1), Chr(&H4E)) > 0 Then
            lng_S = InStr(mstrCode, Chr(&H4E))
        ElseIf InStr(Mid(mstrCode, 1, 1), Chr(&H45)) > 0 Then
            lng_S = InStr(mstrCode, Chr(&H45))
        ElseIf InStr(Mid(mstrCode, 1, 1), Chr(&H43)) > 0 Then
            lng_S = InStr(mstrCode, Chr(&H43))
        End If
        
        If lng_S > 0 Then
           lng_E = InStr(mstrCode, Chr(&HD) & Chr(&HA))
        End If

    Loop
    strResult = Mid(aResult, 3)    '���ؼ�����
    Exit Sub
errH:
    WriteLog "FH400", "FH", Err.Number & " : " & Err.Description
End Sub




