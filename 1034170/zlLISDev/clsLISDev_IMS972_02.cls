VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_IMS972_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:IMS 972 ����ʷ�����
'����:
'����:������
'�û�:̫ԭ
'���ڣ�2012-12-17

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
   
    Dim lng_S As Long, lng_E As Long, str_in As String
    Dim str_Code As String 'һ��������¼
    Dim str_Line As String 'һ������
    Dim str_�걾�� As String, str_��Ŀ As String, str_��� As String, str_���� As String
    Dim str_���� As String, str_SendLine As String
    
    Dim lng_Loop As Long, var_Code As Variant
    strResult = ""
    str_in = strSource
    
    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
    If mstrCode Like "*PAT_ID*" & Chr(&HD) & "*" Then                          '�յ�һ�������ļ�¼,��ʼ����
        WriteLog "���� IMS 972", mstrCode, ""
        Do While InStr(mstrCode, "PAT_ID") > 0
'            lng_S = InStr(mstrCode, "nCa")
             lng_S = 1

            lng_E = InStr(mstrCode, "PAT_ID")
            lng_E = InStr(lng_E, mstrCode, Chr(&HD))
            
            str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
            mstrCode = Mid(mstrCode, lng_E + 1)
            var_Code = Split(str_Code, Chr(&HD))
            str_��Ŀ = ""
            For lng_Loop = LBound(var_Code) To UBound(var_Code)

                str_Line = var_Code(lng_Loop)
                
                If InStr(str_Line, "PAT_ID") > 0 Then                       '��Ŀ ��� ��
                    str_�걾�� = Val(Mid(str_Line, 7))
                ElseIf InStr(",pH,TCa,iCa,nCa,Cl,K,Na,", "," & Trim(Mid(str_Line, 1, 3)) & ",") > 0 Then
                    str_��� = Val(Mid(str_Line, 4))
                    If Val(str_���) <> 0 Then
                        str_��Ŀ = str_��Ŀ & "|" & Trim(Mid(str_Line, 1, 3)) & "|" & str_���
                    End If
                End If

            Next
            If Val(str_�걾��) <> 0 And str_��Ŀ <> "" Then
                strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|" & Val(str_�걾��) & "| |Ѫ��|0" & str_��Ŀ
            End If
        Loop
    End If
    
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
        WriteLog "���� IMS 972", strResult, ""
    End If
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





