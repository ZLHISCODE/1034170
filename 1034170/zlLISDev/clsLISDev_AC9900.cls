VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AC9900"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:AC9900 ȫ�Զ�����ʷ�����
'����:�򽭰µϿ�ҽ�������������ι�˾
'����:���
'�û�:������˾-��ΰ(˼��������ҽԺ)

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
    Dim lng_S As Long, lng_E As Long, str_Tmp As String
    Dim str_Data As String
    Dim varLine As Variant, blnEnd As Boolean
    Dim i As Integer, str_Line As String
    Dim lngLength As Long
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    
    If strSource Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
        blnEnd = False
        strReserved = strSource
        
        Do While Not blnEnd
            lngLength = InStr(strReserved, Chr(3)) - InStr(strReserved, Chr(2)) + 1
            str_Data = Mid(strReserved, InStr(strReserved, Chr(2)), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, Chr(3)) + 1)
            
            str_��Ŀ = ""
            str_ʱ�� = Trim(Mid(str_Data, 2, 20))
            str_������ = Trim(Mid(str_Data, 22, 6))
            
            str_Line = Mid(str_Data, 28, 15)
            If Trim(Mid(str_Line, 1, 6)) <> "" And Trim(Mid(str_Line, 7)) <> "" Then
                str_��Ŀ = str_��Ŀ & "|" & Trim(Mid(str_Line, 1, 6)) & "|" & Trim(Mid(str_Line, 7))
            End If
            varLine = Split(Mid(str_Data, 29), Chr(&H2A))
            
            For i = LBound(varLine) To UBound(varLine) - 1
                str_Line = varLine(i)
                If Trim(Mid(str_Line, 1, 6)) <> "" And Trim(Mid(str_Line, 7)) <> "" Then
                    str_��Ŀ = str_��Ŀ & "|" & Trim(Mid(str_Line, 1, 6)) & "|" & Trim(Mid(str_Line, 7))
                End If
            Next
            If Val(str_������) > 0 And IsDate(str_ʱ��) And str_��Ŀ <> "" Then
                strResult = strResult & "||" & str_ʱ�� & "|" & str_������ & "|" & " |" & "��Һ|0" & str_��Ŀ
            End If
            If Not strReserved Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
                blnEnd = True
            End If
           
        Loop
    Else
       strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
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



