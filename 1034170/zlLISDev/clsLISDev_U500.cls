VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_U500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'�����ж�Ժ

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
    
    Dim iStrPoint As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim iPos As Integer
    Dim strDate As String, lngSampleNO As Long, strItemName As String, strItemValue As String
    Dim strLine As String, blnEnd As Boolean
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    If strSource Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            
            strLine = GetStr_Section(strReserved, Chr(2), Chr(3))
            strReserved = Mid(strReserved, InStr(strReserved, Chr(3)) + 1)
            
            
            iStrPoint = 1: blnReadItem = False
            strDate = "": lngSampleNO = -1
            Do While iStrPoint <= Len(strLine)
                If Not blnReadItem Then '���ұ걾������Ϣ
                    strBuffer = ReadLine(strLine, iStrPoint)
                    Do While Len(strBuffer) > 0
                        Select Case True
                            Case InStr(UCase(strBuffer), UCase("NO.")) > 0 '�걾��
                                iPos = InStr(UCase(strBuffer), UCase("NO."))
                                lngSampleNO = Val(Mid(strBuffer, iPos + 3, 8))
                                strDate = Trim(Mid(strBuffer, iPos + 11, 10))
                                
                                strBuffer = ReadLine(strLine, iStrPoint)
                                strDate = strDate & " " & Trim(Mid(strBuffer, iPos + 13, 10))
                            Case Else
                                strBuffer = ""
                                strItemValue = ""
                        End Select
                    Loop
                    If strDate <> "" And lngSampleNO > 0 Then '������Ϣ������ɣ���λ��������Ŀ��ʼλ��
                        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                        strData = strDate
                        '������ţ�����Ϊ����
                        strData = strData & "|" & lngSampleNO
                        '������
                        strData = strData & "| " '��������֧��
                        '�걾
                        strData = strData & "|ѪҺ" '��������֧��
                        '�ʿ�Ʒ
                        strData = strData & "|0" '��������֧�֣���Ϊ���ʿ�Ʒ
                        
                        Do While iStrPoint <= Len(strLine)
                            strBuffer = Trim(ReadLine(strLine, iStrPoint))
                            If strBuffer = "" Then Exit Do
                        Loop
                        blnReadItem = True
                    End If
                Else '������Ŀ
                    strBuffer = ReadLine(strLine, iStrPoint)
                    Do While Len(strBuffer) > 0
                        iPos = 1
                        strItemName = UCase(Trim(Mid(strBuffer, 1, 4)))
                           
                        If InStr("WBC,KET,NIT,URO,BIL,PRO,GLU,SG,BLD,PH,VC", Replace(strItemName, "*", "")) > 0 And strItemName <> "" Then
                            If Trim(Mid(strBuffer, 12)) = "Normal" Then
                                strItemValue = strItemValue & "|" & strItemName
                                strItemValue = strItemValue & "|Normal"
    '                        ElseIf InStr(strItemName, "*") > 0 Then
    '                            strData = strData & "|" & strItemName
    '                            strData = strData & "|" & Val(Mid(strBuffer, 12, 5)) & Trim(Mid(strBuffer, 5, 7))
    '                        'ElseIf Trim(Mid(strBuffer, 5, 8)) <> "-" Then
                            Else
                                strItemValue = strItemValue & "|" & strItemName
                                strItemValue = strItemValue & "|" & IIf(Trim(Mid(strBuffer, 12, 5)) = "", _
                                            "" & Trim(Trim(Mid(strBuffer, 5, 7))), _
                                            Trim(Replace(Replace(Replace(Replace(Mid(strBuffer, 12), "Cell/uL", ""), "mmol/L", ""), "umol/L", ""), "g/L", "")) & _
                                            Replace("(" & Trim(Trim(Mid(strBuffer, 5, 7))) & ")", "()", ""))
                                
                            End If
                        End If
                        strBuffer = ReadLine(strLine, iStrPoint)
                        If Mid(strBuffer, 1, 1) = "" Then Exit Do
                    Loop
                End If

            Loop
            If strItemValue <> "" Then
                strResult = strResult & "||" & strData & strItemValue
            End If
'            '����δ������ԭʼ��
'            If iCurrGrpEntry > 0 Then
'                strReserved = Mid(strSource, iCurrGrpEntry)
'            Else
'                '���豸���͡������յ�������
'                strCmd = Chr(&H6&)
'            End If
            
            If Not strReserved Like "*" & Chr(&H2) & "*" & Chr(&H3) & "*" Then
                blnEnd = True
            End If
        Loop
    Else
        strReserved = strSource
        strCmd = Chr(&H6&)
    End If
    
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Replace(Replace(Replace(Replace(Mid(strResult, 3), "*", ""), "()", ""), "|.", "|0."), "+-", "��")
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

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    'ָ�붨λ����һ������֮��
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = Chr(10) Then
            iPoint = iPoint + 1
            Exit Do
        End If
        iPoint = iPoint + 1
    Loop
End Sub

Private Function ReadSubString(ByVal strSource As String, iPoint As Integer, Optional ByVal Length As Integer = 0) As String
    '��ȡ�Ӵ���ͬʱָ���ƶ�����ȡ���ַ���֮��
    On Error Resume Next
    If Length = 0 Then
        ReadSubString = Mid(strSource, iPoint)
        iPoint = Len(strSource) + 1
    Else
        ReadSubString = Mid(strSource, iPoint, Length)
        iPoint = iPoint + Length
    End If
End Function

Private Function ReadLine(ByVal strSource As String, iPoint As Integer) As String
    '��ȡһ���ַ��������ַ�ָ���Ƶ���һ�п�ʼλ��
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, Chr(13) + Chr(10))
    If iPos = 0 Then iPos = Len(strSource) + 1
    
    ReadLine = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 2
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '��ȡһ��������Ŀ���ƣ������ַ�ָ���Ƶ���Ŀ����֮��
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function
