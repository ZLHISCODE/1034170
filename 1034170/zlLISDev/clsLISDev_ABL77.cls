VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ABL77"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Ѫ�������� ABL77
'����:�����׶�
'����:�¶�
'�û�:��ҽ��Ժ

'pH��pCO2 ��pO2��HCT�� K+��Na+��Ca2+ ��Cl-
Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = Chr(6)
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
    
    Dim varLine As Variant, varItem As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    If strSource Like "*" & Chr(&H5&) & "*" & Chr(&H4&) & "*" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
            strLine = GetStr_Section(strReserved, Chr(&H5&), Chr(&H4&))
            strReserved = Mid(strReserved, InStr(strReserved, Chr(&H4&)) + 1)
            
            str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = "": strTmp = ""
            '����
            varLine = Split(strLine, vbCrLf)
            strItem = ""
            For iRow = LBound(varLine) To UBound(varLine)
                If iRow = 1 Then
                    '��ʼ
                    If InStr(varLine(iRow), "|") > 0 Then
                        If UBound(Split(varLine(iRow), "|")) > 4 Then
                            str�걾�� = Val(Split(varLine(iRow), "|")(4))
                        End If
                    End If
                    
                ElseIf iRow = 3 Then
                    If InStr(varLine(iRow), "|") > 0 Then
                        If UBound(Split(varLine(iRow), "|")) > 14 Then
                            strTmp = Split(varLine(iRow), "|")(14)
                            str���� = Format(Mid(strTmp, 1, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2), "yyyy-MM-dd")
                            strTmp = Trim(Split(varLine(iRow), "|")(5))
                            If strTmp <> "" Then
                                strItem = "|" & GetStr_Section(Split(varLine(iRow), "|")(3), "^", "^") & "|" & strTmp
                            End If
                        End If
                    End If
                ElseIf iRow > 3 Then
                    If InStr(varLine(iRow), "|") > 0 Then
                        If UBound(Split(varLine(iRow), "|")) > 5 Then
                            strTmp = Trim(Split(varLine(iRow), "|")(5))
                            If strTmp <> "" Then
                                strItem = strItem & "|" & GetStr_Section(Split(varLine(iRow), "|")(3), "^", "^") & "|" & strTmp
                            End If
                        End If
                    End If
                End If
            Next
            
            If strItem <> "" And Val(str�걾��) <> 0 Then
                strResult = strResult & "||" & str���� & "|" & str�걾�� & "| |ѪҺ|0" & Replace(strItem, "^", "")
            End If
            If Not strReserved Like "*" & Chr(&H5&) & "*" & Chr(&H4&) & "*" Then
                blnEnd = True
            End If
        Loop
    Else
        '�յ�����Ϣ��ȫ,����,���´��ٽ���
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    strCmd = Chr(6)
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
