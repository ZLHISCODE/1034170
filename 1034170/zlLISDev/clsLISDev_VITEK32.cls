VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_VITEK32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  ΢���� VITEK32
'ҽԺ:  ��̨ɽҽԺ
'����:  �¶�
'ʱ��:  2008/08/19

Dim mstrCode  As String

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
    '   ΢��������ʽ ͨ����|�������^��ʽ^���ֵ
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����

    
    Dim aData() As String
    Dim intloop As Integer
    Dim strDate As String               '����
    Dim strSampleNO As String           '�걾��
    Dim strLine As String               '��ǰҪ�������ִ�
    Dim lngBegin As Long                '��ʼλ��
    Dim lngEnd As String                '����λ��
    
    Dim strϸ�� As String, strҩ�� As String
    
    On Error GoTo errH
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    'û���ҵ���ʼλ��ʱ�˳�
    
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
    
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H3)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H4)) > 0 Then
        strCmd = Chr(&H6)
    End If
    
    'û���ҵ�����λ��ʱ�˳�
    lngBegin = InStr(mstrCode, Chr(&H5))
    lngEnd = InStr(mstrCode, Chr(&H4))
     
    Do While lngBegin > 0 And lngEnd > 0
    
        'ȡ��ǰһ�εĽ���������ķ���
        strLine = Mid(mstrCode, lngBegin, lngEnd - lngBegin)
        mstrCode = Mid(mstrCode, lngEnd + 1)
        
        '��ʼ����
        aData = Split(strLine, "|")
        '����
        strDate = Mid(aData(1), 3, 4) & "-" & Mid(aData(1), 7, 2) & "-" & Mid(aData(1), 9, 2)
        '�걾��
        strSampleNO = Mid(aData(3), 3)
        If Val(strSampleNO) > 0 Then
            strϸ�� = "": strҩ�� = ""
            
            '��ʼ�������
            For intloop = 4 To UBound(aData)
                If Left(aData(intloop), 2) = "a1" Then
                    '�򷵻���������ҩ��������Ϣ,���Ծ�����Ϊ   MIC
                    strҩ�� = strҩ�� & "|" & Mid(aData(intloop), 3) & "|" & Mid(aData(intloop + 3), 3) & "^1^" & Mid(aData(intloop + 2), 3)
                ElseIf Left(aData(intloop), 2) = "o1" Then
                    strϸ�� = Mid(aData(intloop), 3)
                End If
            Next
            If strϸ�� <> "" And strҩ�� <> "" Then
                strResult = strResult & "||" & strDate & "|" & Val(strSampleNO) & "| |" & strϸ�� & "|0" & strҩ��
            End If
        End If
        lngBegin = InStr(mstrCode, Chr(&H5))
        lngEnd = InStr(mstrCode, Chr(&H4))
    Loop
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




