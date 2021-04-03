VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CD3200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Cell Dyn 3200
'����:��������
'����:�¶�
'�û�:����ʡ����ҽԺ

'Data

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
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    Dim varline As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String
    Dim lngLength As Long
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    
    If strSource Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
            lngLength = InStr(strReserved, Chr(3)) - InStr(strReserved, Chr(2)) + 1
            strLine = Mid(strReserved, InStr(strReserved, Chr(2)), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, Chr(3)) + 1)
            varline = Split(strLine, ",")
            
            str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = "": strTmp = ""
            If UBound(varline) >= 73 Then
                If varline(0) Like "*" & Chr(2) & """   ""*" Then
                    '��������
                    strTmp = Replace(varline(13), """", "")
                    str���� = Format(CDate(Right(strTmp, 2) & "-" & Replace(Left(strTmp, 5), "/", "-")), "yyyy-MM-dd")
                    strTmp = Replace(varline(14), """", "")
                    str���� = str���� & " " & strTmp
                    
                    'str�걾�� = Replace(varline(6), """", "")
                    str�걾�� = Replace(varline(3), """", "")
                    If str�걾�� Like "*R*T*" Then
'                        strTmp = Mid(str�걾��, InStr(str�걾��, "R") + 1, InStr(str�걾��, "T") - InStr(str�걾��, "R"))
'                        str�걾�� = Val(CStr(Val(strTmp) - 1) & Mid(str�걾��, InStr(str�걾��, "T") + 1))
                        strTmp = Val(Mid(str�걾��, InStr(str�걾��, "T") + 1))
                        str�걾�� = (Val(Mid(str�걾��, 2)) - 1) * 10 + Val(strTmp)
                    Else
                        str�걾�� = Val(str�걾��)
                    End If
                    str�����˱걾�ʿ� = " |ѪҺ|" & IIf(str�걾�� >= 2000, 1, 0)
                    
                    For iRow = 18 To 39
                        If Val(varline(iRow)) <> 0 Then
                            strItem = strItem & "|" & iRow & "|" & Val(varline(iRow))
                        End If
                    Next
                    
                    If strItem <> "" And Val(str�걾��) <> 0 Then
                        strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
                    End If
                    
                ElseIf varline(0) Like "*" & Chr(2) & """PLT""*" Then
                    'PLTͼ�� �ݲ�����
                ElseIf varline(0) Like "*" & Chr(2) & """RBC""*" Then
                    'RBCͼ�� �ݲ�����
                ElseIf varline(0) Like "*" & Chr(2) & """WBC""*" Then
                    'WBCͼ�� �ݲ�����
                End If
            End If
            If Not strReserved Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
                blnEnd = True
                'strReserved = Replace(strReserved, "^", "#") ' ��^��ԭΪ#
            End If
        Loop
    Else
        '�յ�����Ϣ��ȫ,����,���´��ٽ���
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If

End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:������Ҫ���������͵ı걾��Ϣ
    '           �걾֮����||�ָ�
    '           Ԫ��֮����|�ָ�
    '           ��0��Ԫ�أ�����ʱ��
    '           ��1��Ԫ�أ��������
    '           ��2��Ԫ�أ�������
    '           ��3��Ԫ�أ��걾
    '           ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '           ��5��9Ԫ�أ�ϵͳ����
    '           �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    '   iSendStep  : ����
    '   blnSuccess : �Ƿ���ʧ��
    '   strResponse: ������Ӧ����Ϣ
    '   blnUndo    : �Ƿ�ȡ��.
    '
    '   ����ֵ�������������ִ���ָ�
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function
