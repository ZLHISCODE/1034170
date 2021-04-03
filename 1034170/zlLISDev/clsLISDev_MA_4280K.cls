VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MA_4280K"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:MA-4280K
'����:���ֻ�ͨ
'����:����
'�û�:�����ж�Ժ

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
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    '�������������д�������������ע�����벻Ҫɾ����
    Dim str_ʱ�� As String, str_������ As String, str_��Ŀ As String, str_��� As String
    Dim lon_str As Long, lon_end As Long
    Dim str_Data As String
    Dim varLine As Variant
    Dim i As Integer, str_Line As String
    
    
    
    lon_str = InStr(strSource, Chr(&H3F))
    lon_end = InStr(strSource, Chr(&H3))
    If lon_str > 0 And lon_end > 0 Then
    
        str_Data = Mid(strSource, lon_str, lon_end)
        str_��Ŀ = ""
        
        str_ʱ�� = Format(Now, "yyyy-MM-dd hh:MM:ss") 'Format(Mid(str_data, 13, 4) & "-" & Mid(str_data, 18, 2) & "-" & Mid(str_data, 22, 2) & " " & Mid(str_data, 37, 2) & ":" & Mid(str_data, 40, 2) & ":" & Mid(str_data, 43, 2), "YYYY-MM-DD hh:mm:ss")
        varline = Split(str_Data, Chr(&HD))
        For i = LBound(varline) To UBound(varline)
            str_Line = varline(i)
            If str_Line Like "*NO.*" Then
                str_������ = Val(Mid(str_Line, InStr(str_Line, "NO.") + 3))
            ElseIf InStr(str_Line, "") <= 1 Then
                str_��Ŀ = str_��Ŀ & "|" & Mid(str_Line, 1, 4) & "|" & Trim(Mid(str_Line, 5, 7))
            End If
            
        Next
        
        strResult = str_ʱ�� & "|" & str_������ & "|" & " |" & "��Һ" & "|0" & str_��Ŀ
    Else
        strReserved = strSource
    End If

End Sub


Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
End Function