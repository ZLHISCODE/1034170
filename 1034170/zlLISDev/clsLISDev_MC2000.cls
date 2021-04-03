VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MC2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:MC-2000 ��ͨ��Ѫ����
'����:��������
'����:�¶�
'�û�:����202ҽԺ
Private mstrCode As String
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
    Dim str_Line As String, lng_S As Long, lng_E As Long, var_line As Variant, str_Item As String
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    lng_E = InStr(mstrCode, Chr(&HD) & Chr(&HA))
    If lng_E > 0 Then
        lng_S = lng_E - 57
        If lng_S > 0 Then
            str_Line = Mid(mstrCode, lng_S, lng_E - lng_S)
            mstrCode = Mid(mstrCode, lng_E + 2)
        
            var_line = Split(str_Line, Chr(&H9))
            If UBound(var_line) >= 10 Then
                str_ʱ�� = Format(Now, "yyyy-MM-dd")
                str_������ = var_line(0)
                
                If Val(str_������) > 0 Then
                    str_Item = var_line(1)
                    Select Case str_Item
                    Case "01"
                        strResult = str_ʱ�� & "|" & Val(str_������) & "| |ѪҺ|0|PT|" & Val(var_line(5)) / 10 & "|INR|" & Val(var_line(7)) / 100
                    Case "02"
                        strResult = str_ʱ�� & "|" & Val(str_������) & "| |ѪҺ|0|APTT|" & Val(var_line(5)) / 10
                    Case "03"
                        strResult = str_ʱ�� & "|" & Val(str_������) & "| |ѪҺ|0|TT|" & Val(var_line(5)) / 10
                    Case "04"
                        strResult = str_ʱ�� & "|" & Val(str_������) & "| |ѪҺ|0|FIB|" & Val(var_line(10)) / 10
                    End Select
                End If
            End If
        Else
            mstrCode = Mid(mstrCode, lng_E + 2)
        End If
        
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

