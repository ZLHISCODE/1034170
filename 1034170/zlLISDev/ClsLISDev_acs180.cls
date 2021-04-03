VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_ACS180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ACS180 ��ѧ�������߷�����
'����:�ݶ�
'����:�¶�
'�û�:����ҽԺ

Private mstrCode As String
Private mstr�걾�� As String
Private mstr��Ŀ As String

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
    Dim str_���� As String
    Dim lng_Loop As Long
    Dim str_All As String
    strResult = ""
    str_in = strSource
    '-------------------------------------------------------------------------------------------------------------
    If InStr(str_in, Chr(&H5)) > 0 Then                                         '�յ���ʼ�� 05 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    If InStr(str_in, Chr(&H3)) > 0 Then                                         '�յ��н����� 03 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    
    '-------------------------------------------------------------------------------------------------------------
    
    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    lng_S = InStr(mstrCode, Chr(&H2))
    lng_E = InStr(mstrCode, Chr(&H3))
    
    If lng_S > 0 And lng_E > 0 Then WriteLog "���� acs180", mstrCode, ""
    Do While lng_S > 0 And lng_E > 0                       '�յ�һ�������ļ�¼,��ʼ����
        
        
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)

            Do While str_Code <> ""
                lng_S = InStr(str_Code, Chr(&H2))
                lng_E = InStr(str_Code, Chr(&H3))
                
                If lng_S > 0 And lng_E > 0 Then
                    str_Line = Mid(str_Code, lng_S, lng_E - lng_S + 1)
                    str_Code = Mid(str_Code, lng_E + 1)
                    
                    If str_Line Like "*" & Chr(&H2) & "?O|*" Then                           '�걾��,��Ŀ ��
                        mstr�걾�� = Val(Split(str_Line, "|")(2))
                        If mstr�걾�� = 0 Then mstr�걾�� = ""
                        mstr��Ŀ = Split(str_Line, "|")(4)
                        If mstr��Ŀ = "^^^Cpsa-Devision" Then mstr��Ŀ = "^^^CpsaDevision"
                    ElseIf str_Line Like "*" & Chr(&H2) & "?R|*" Then                          '��Ŀ ��� ��
                        If mstr�걾�� <> "" And mstr��Ŀ <> "" Then
                            If mstr��Ŀ = Split(str_Line, "|")(2) Then
                                str_���� = Split(str_Line, "|")(12)
                                str_��� = Split(str_Line, "|")(3)
                                If Trim(str_���) = ">>>" Then str_��� = "������"
                                If Trim(str_���) = "<<<" Then str_��� = "������"
                                
                                If str_���� <> "" Then
                                    str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, 2)
                                    strResult = strResult & "||" & str_���� & "|" & mstr�걾�� & "| |Ѫ��|0|" & Replace(mstr��Ŀ, "^^^", "") & "|" & str_���
                                    
                                End If
                            End If
                        End If
                        mstr�걾�� = "": mstr��Ŀ = ""
                        
                    End If
                Else
                    str_Code = ""
                End If
            Loop
        lng_S = InStr(mstrCode, Chr(&H2))
        lng_E = InStr(mstrCode, Chr(&H3))
    Loop
    
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
        WriteLog "���� acs180 End", strResult, ""
    End If
End Sub

Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    SendSample = ""
End Function



