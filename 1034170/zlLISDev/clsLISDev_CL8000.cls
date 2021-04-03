VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CL8000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:���� CL-8000  ����������
'����:�ձ�����
'����:�¶�
'�û�:�ɶ�XXҽԺ���»Բɵ����ݣ�
'�޸���:���
'�޸�����:2010-01-29

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
    Dim str_�걾�� As String, str_��Ŀ As String, str_���� As String, str_�ʿ� As String, str_��� As String
    Dim str_Tmp As String
    Dim strSmpStart As String
    
    strResult = ""
    str_in = strSource
    '-------------------------------------------------------------------------------------------------------------
    '��Ӧ��
    '-------------------------------------------------------------------------------------------------------------
    strSmpStart = ReadIni("CL8000", "�걾����ʼλ��", App.Path & "\CL8000.ini")
    If strSmpStart = "" Then strSmpStart = "7": Call WriteIni("CL8000", "�걾����ʼλ��", "7", App.Path & "\CL8000.ini")
    
    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
    Do While mstrCode Like "*" & Chr(&H2) & "*" & Chr(&H3) & "*"                         '�յ�һ�������ļ�¼,��ʼ����
    
        lng_S = InStrRev(mstrCode, Chr(&H2))
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        str_�ʿ� = Mid(str_Code, 2, 1)
        
        If str_�ʿ� = "Q" Then
            str_�ʿ� = "1"
        ElseIf str_�ʿ� = "D" Then
            str_�ʿ� = "0"
        Else
            GoTo errFormat
        End If
        
        If str_�ʿ� = "0" Then
            str_�걾�� = Val(Mid(str_Code, Val(strSmpStart), 4))
            str_���� = Format(Now, "yyyy-MM-dd HH:MM")
        
            str_Line = Mid(str_Code, 26)
            Do While Len(str_Line) > 22
                str_Tmp = Left(str_Line, 22)
                str_Line = Mid(str_Line, 23)
                str_��� = Trim(Mid(str_Tmp, 7, 8))
                If str_��� <> "" Then
                    str_��Ŀ = str_��Ŀ & "|" & Trim(Mid(str_Tmp, 2, 3)) & "|" & Val(str_���) / 1000
                End If
            Loop
            
            If Val(str_�걾��) <> 0 And str_��Ŀ <> "" Then
                strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|0" & str_��Ŀ
            End If
        ElseIf str_�ʿ� = "1" Then
            str_�걾�� = Mid(str_Code, 9, 3)
            str_���� = Format(Now, "yyyy-MM-dd HH:MM")
        
            str_Line = Mid(str_Code, 19)
            Do While Len(str_Line) > 24
                str_Tmp = Left(str_Line, 24)
                str_Line = Mid(str_Line, 25)
                str_��� = Trim(Mid(str_Tmp, 5, 8))
                If str_��� <> "" Then
                    str_��Ŀ = str_��Ŀ & "|" & Trim(Mid(str_Tmp, 2, 3)) & "|" & Val(str_���) / 1000
                End If
            Loop
            
            If Val(str_�걾��) <> 0 And str_��Ŀ <> "" Then
                strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|1" & str_��Ŀ
            End If

        End If
    Loop

    
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
        
    End If
    Exit Sub
errFormat:
    
    
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