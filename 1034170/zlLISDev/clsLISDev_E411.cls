VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_E411"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'����:�绯ѧ���� E411
'����:Roche ����
'����:�¶�
'�û�:����ҽԺ
'�޸���:���
'�޸�����:2009-08-28

Private mstrCode As String    '��������

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
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
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    Dim lng_S As Long, lng_E As Long, i_Count As Integer
    Dim str_in As String, str_Line As String, str_���� As String, var_line As Variant
    Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String, str_��� As String
    Dim str_WaitSendLine As String, str_���� As String, str_�ʿ� As String
    Dim str_Last�걾�� As String
    str_in = strSource
        
    If InStr(str_in, Chr(&HA)) > 0 Then                 '�յ�ETX(03)
        strCmd = Chr(&H6)
    End If
    
    If InStr(str_in, Chr(&H5)) > 0 And strCmd = "" Then               '�յ�ENQ(05) Ӧ�� ETX(03)
        strCmd = Chr(&H6)
    End If

    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
    lng_S = InStr(mstrCode, Chr(&H5))
    lng_E = InStr(mstrCode, Chr(&H4))
    If lng_S > 0 And lng_E > 0 Then
        
        str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        WriteLog "��ʼ", "�յ���������", str_Line
        
        If str_Line Like "*" & Chr(&H2) & "?R|*" Then     '������
            var_line = Split(str_Line, Chr(&HA))
            
            For i_Count = LBound(var_line) To UBound(var_line)
                str_Line = var_line(i_Count)
                If str_Line Like Chr(&H2) & "?O|*" Then
                    str_���� = Split(str_Line, "|")(6)
                    str_���� = Left(str_����, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2)
                    
                    'str_�걾�� = Mid(Split(str_Line, "|")(3), 1, InStr(Split(str_Line, "|")(3), "^") - 1)
                    str_�걾�� = Replace(Split(str_Line, "|")(2), "@", "")
                    str_���� = Trim(Split(str_Line, "|")(5))
                    
                    str_�ʿ� = "0"
                    If Trim(Split(str_Line, "|")(11)) = "Q" Then str_�ʿ� = "1"
                    If Trim(Split(str_Line, "|")(11)) = "X\Q" Then str_�ʿ� = "1"
                    
                    If Val(Trim(str_�걾��)) > 0 Then
                        If Len(Trim(str_�걾��)) < 12 Then
                            '�걾��
                            str_�걾�� = Trim(str_�걾��) & IIf(str_���� = "S", "^1^", "")
                        Else
                            '����
                            str_�걾�� = "^" & IIf(str_���� = "S", "1^", "^") & Trim(str_�걾��)
                        End If
                    Else
                        str_�걾�� = ""
                    End If
                    
                    str_��Ŀ = ""
                ElseIf str_Line Like Chr(&H2) & "?R|*" Then
                    str_��� = Trim(Split(str_Line, "|")(3))
                    If str_��� <> "" Then
                        If InStr(str_���, "^") > 0 Then
                            str_��Ŀ = "|" & Split(Split(str_Line, "|")(2), "^")(3) & "|" & Split(str_���, "^")(1)
                        Else
                            str_��Ŀ = "|" & Split(Split(str_Line, "|")(2), "^")(3) & "|" & str_���
                        End If
                        If IsDate(CDate(str_����)) And str_�걾�� <> "" And str_��Ŀ <> "" Then
                            If str_Last�걾�� = str_�걾�� Then
                                strResult = strResult & str_��Ŀ
                            Else
                                strResult = strResult & "||" & str_���� & "|" & str_�걾�� & "| |Ѫ��|" & str_�ʿ� & str_��Ŀ
                                str_Last�걾�� = str_�걾��
                            End If
                        End If
                    End If
                End If
            Next
            If strResult <> "" Then strResult = Mid(strResult, 3)
            WriteLog "����", "������", strResult
        End If
    End If
    
End Sub


Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    
    'iSendStep����ǰ���Ͳ����Ĳ��衣�����������ͽ������򷵻�0�������Զ���1���ء�
    'strResponse����ǰ��������Ӧ����
    'strSampleInfo�����͵ı걾��Ϣ��
    '   �걾֮����||�ָ�
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ���
    '   ��5�����أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��

    Dim var_�걾 As Variant, i_Count As Integer, i_Loop As Integer, str_��¼ As String, str_�걾�� As String
    Dim i_���� As Integer, str_���� As String, str_������Ŀ As String, str_�걾���� As String
    Dim str_SendItem As String, str_Line As String
    
    SendSample = ""

    
End Function




