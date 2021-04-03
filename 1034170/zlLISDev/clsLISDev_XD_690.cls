VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XD_690"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  XD690 ����� Ѹ��
'ҽԺ:  �����е�һ����ҽԺ
'����:  �ƽ���
'ʱ��:  2013-06-20

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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)
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
    '   ΢��������ʽ ͨ����|�������^��ʽ^���ֵ
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����

    Dim str_Code As String
    Dim str_Line As String               '��ǰҪ�������ִ�
    Dim lng_S As Long, lng_E As String, i As Integer               '����λ��
    
    Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String
    Dim varLine As Variant
    mstrCode = mstrCode & strSource
    On Error GoTo errH
    strResult = ""
    If mstrCode Like "*XD  690*" & Chr(&HA) & "*CL*" Then
        
        Do While InStr(mstrCode, "CL") > 0
            lng_S = InStr(mstrCode, "XD  690")
            lng_E = InStr(lng_S, mstrCode, "CL")
            lng_E = InStr(lng_E, mstrCode, "/L")
            str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
            mstrCode = Mid(mstrCode, lng_E + 2)
            
            str_���� = "": str_�걾�� = "": str_��Ŀ = ""
            varLine = Split(str_Code, Chr(&HA))
            For i = 0 To UBound(varLine)
                str_Line = varLine(i)
                If str_Line Like "DATE*" Then
                    str_���� = Trim(Mid(str_Line, 5))
                ElseIf str_Line Like "TIME*" Then
                    str_���� = str_���� & " " & Trim(Replace(str_Line, "TIME", ""))
                    str_���� = Format(str_����, "yyyy-mm-dd hh:mm:ss")
                ElseIf str_Line Like "Sample No*" Then
                    str_�걾�� = Val(Mid(str_Line, 12))
                ElseIf str_Line Like "*=*" Then
                    str_��Ŀ = str_��Ŀ & "|" & Trim(Mid(str_Line, 1, InStr(str_Line, "=") - 1)) & "|" & Trim(Mid(str_Line, InStr(str_Line, "=") + 1, 5))
                End If
                
            Next
            
            If Val(str_�걾��) <> 0 And IsDate(str_����) And str_��Ŀ <> "" Then
                strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |ѪҺ|0" & str_��Ŀ
            End If
        Loop
    End If
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errH:
    WriteLog "err", Err.Description, strResult
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    
End Function

