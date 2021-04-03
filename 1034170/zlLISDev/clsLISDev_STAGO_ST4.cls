VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_STAGO_ST4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:STAGO -ST4��Ѫ��
'����:���� STAGO
'����: �¶�
'�û�:���콭��ҽԺ


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
    '   ÿ���������||�ָ�,Ԫ��֮���� �ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    Dim iStrPoint As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim varIn As Variant '�����������
    Dim strTime As String
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim strLine As String, intCount As Integer, i As Integer
    Dim str�걾�� As String, str���� As String, str�����˱걾�ʿ� As String, strָ�� As String
    Dim strIn As String
    On Error GoTo errHandle
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    If strSource = "" Then Exit Sub
    
    If InStr(strSource, Chr(3)) = 0 Or InStr(strSource, Chr(2)) = 0 Then
        strReserved = strSource
        Exit Sub
    End If
    
    strָ�� = Trim(Mid(strSource, 11, 2))
    If strָ�� = "1" Then strָ�� = "PT"
    If strָ�� = "2" Then strָ�� = "APTT"
    If strָ�� = "" Then Exit Sub
    str�����˱걾�ʿ� = " |ѪҺ|0"
    
    'strָ�� = strָ�� & " " & Trim(Mid(strSource, 13, 5))
    WriteLog "STAGO ST4", "׼������", strSource
    
    strSource = Replace(strSource, "}", "")
    If InStr(strSource, Chr(3)) > 0 And InStr(strSource, Chr(2)) > 0 Then
        'һ���Դ�����һ�ν��
        strIn = strSource
        For iStrPoint = 1 To Len(strSource)
            
            If Mid(strSource, iStrPoint, 1) = Chr(2) Then
                strIn = Mid(strSource, iStrPoint, 1)
            Else
                strIn = strIn & Mid(strSource, iStrPoint, 1)
            End If
            
            If Mid(strSource, iStrPoint, 1) = Chr(3) Then
                strReserved = Mid(strSource, iStrPoint + 1)
                Exit For
            End If
        Next
    Else
        'û����,������,ֱ�ӷ���
        strReserved = strSource
        Exit Sub
    End If
    
    If InStr(strIn, Chr(2)) > 0 And InStr(strIn, Chr(3)) > 0 Then
        iStrPoint = 0
        varIn = Split(strIn, Chr(3))
        For iStrPoint = LBound(varIn) To UBound(varIn) - 1
            intCount = Val(Trim(Mid(varIn(iStrPoint), 26, 2)))
            If intCount <= 0 Then Exit Sub
            str���� = Format(Now, "yyyy-MM-dd HH:mm:ss")
            str�걾�� = Trim(Mid(varIn(iStrPoint), 9, 2))
            If Len(varIn(iStrPoint)) < (intCount * 37 + 27) Then Exit Sub
            
            For i = 1 To intCount
                strLine = Mid(varIn(iStrPoint), 28 + (i - 1) * 37, 37)
                strResult = strResult & "||" & str���� & "|" & str�걾�� & Trim(Mid(strLine, 1, 12)) & "|" & str�����˱걾�ʿ� & "|" & _
                              strָ�� & "|" & IIf(Trim(Mid(strLine, 14, 8)) = "", "����", Trim(Mid(strLine, 14, 8)))
                If strָ�� <> "APTT" Then
                    strResult = strResult & "|" & Trim(Mid(varIn(iStrPoint), 13, 5)) & "|" & IIf(Trim(Mid(strLine, 22, 8)) = "", "����", Trim(Mid(strLine, 22, 8))) & "|" & _
                                  "INR|" & IIf(Trim(Mid(strLine, 30, 8)) = "", "����", Trim(Mid(strLine, 30, 8)))
                End If
            Next
        Next
        If strResult <> "" Then strResult = Mid(strResult, 3)
    End If
    Exit Sub
    
errHandle:
    WriteLog "STAGO ST4", "�������", Err.Description
    'Resume
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

