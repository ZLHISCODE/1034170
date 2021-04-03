VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Fully"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:BSI Fully ȫ�Զ�����������
'����:
'����:�¶�
'�û�:��������������ҽԺ

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
        
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������^�Ƿ���^�����  ����ֻ���������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
   
    Dim lng_S As Long, lng_E As Long
    Dim str_Code As String 'һ��������¼
    Dim str_�걾�� As String, str_��Ŀ As String, str_��� As String, str_���� As String
    Dim str_���� As String

    strResult = ""
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, "R")
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&HD))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '----------------------------------
        If Len(str_Code) >= 154 Then
            str_���� = Mid(str_Code, 147, 8)
            If IsNumeric(str_����) Then
                str_���� = Mid(str_����, 5, 4) & "-" & Mid(str_����, 3, 2) & "-" & Mid(str_����, 1, 2)
                str_�걾�� = Trim(Mid(str_Code, 3, 5))
                If Val(str_�걾��) > 0 Then
                    str_��Ŀ = Trim(Mid(str_Code, 57, 6))
                    str_��� = Trim(Mid(str_Code, 63, 18))
                    If str_��� <> "" And str_��Ŀ <> "" Then
                        strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|0|" & str_��Ŀ & "|" & str_���
                    End If
                End If
            Else
                str_���� = ""
            End If
        End If
        '----------------------------------
        lng_S = InStr(mstrCode, "R")
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&HD))
        End If
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)

End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    
    SendSample = ""
    
End Function






