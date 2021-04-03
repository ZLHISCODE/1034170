VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_UA_66"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:UA-66
'����:
'����:�¶�
'�û�:�ɶ���ͯר��ҽԺ(���Ǹ�)

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    'GetStartCmd = Chr(6)
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
    '   ͼ�����ݷ��ڼ��������ûس����з�(vbcrlf)�ָ�
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    Dim varLine As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String, lngLength As Long, iReturn As Integer
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    If InStr(strSource, vbCrLf) <= 0 Then
        strReserved = strSource
        Exit Sub
    End If
    
    varLine = Split(strSource, vbCrLf)
    strItem = ""
    For iRow = LBound(varLine) To UBound(varLine)
        If IsDate(varLine(iRow)) Then
            If UBound(varLine) > iRow + 14 Then
                str���� = varLine(iRow)
                str�걾�� = Val(varLine(iRow + 2))
                str�����˱걾�ʿ� = " |��Һ|0"
                For iReturn = iRow + 4 To iRow + 14
                    'If InStr(Mid(varLine(iReturn), InStr(varLine(iReturn), vbTab) + 1), "-") <= 0 Then
                    str��Ŀ = Trim(Mid(varLine(iReturn), 1, InStr(varLine(iReturn), vbTab) - 1))
                    strValues = Trim(Mid(varLine(iReturn), InStr(varLine(iReturn), vbTab) + 1))
                    
                    If UCase(strValues) = "NEG." Then
                        strValues = "-"
                    ElseIf UCase(strValues) = "NORM" Then
                        strValues = "norm"
                    ElseIf strValues = "-" Then
                        strValues = "-"
                    Else
                        Select Case UCase(str��Ŀ)
                        Case "LEU"
                            If Val(strValues) <= 15 Then
                                strValues = "�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 15 And Val(strValues) <= 70 Then
                                strValues = "1�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 70 And Val(strValues) <= 125 Then
                                strValues = "2�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 125 And Val(strValues) <= 500 Then
                                strValues = "3�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 500 Then
                                strValues = ">3�� " & Split(strValues, " ")(0)
                            End If
                        Case "URO"
                            If Val(strValues) <= 2 Then
                                strValues = "norm " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 2 And Val(strValues) <= 4 Then
                                strValues = "1�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 4 And Val(strValues) <= 8 Then
                                strValues = "2�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 8 And Val(strValues) <= 12 Then
                                strValues = "3�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 12 Then
                                strValues = "4�� " & Split(strValues, " ")(0)
                            End If
                        Case "BLD"
                            If Val(strValues) <= 0.06 Then
                                strValues = "1�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 0.06 And Val(strValues) <= 0.15 Then
                                strValues = "2�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 0.15 And Val(strValues) <= 0.75 Then
                                strValues = "3�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 0.75 Then
                                strValues = ">3�� " & Split(strValues, " ")(0)
                            End If
                        Case "BIL"
                            If Val(strValues) <= 0.5 Then
                                strValues = "1�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 0.5 And Val(strValues) <= 1 Then
                                strValues = "2�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 1 And Val(strValues) <= 2 Then
                                strValues = "3�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 2 Then
                                strValues = ">3�� " & Split(strValues, " ")(0)
                            End If
                        Case "KET"
                            If Val(strValues) <= 15 Then
                                strValues = "1�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 15 And Val(strValues) <= 30 Then
                                strValues = "2�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 30 And Val(strValues) <= 80 Then
                                strValues = "3�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 80 Then
                                strValues = ">3�� " & Split(strValues, " ")(0)
                            End If
                        Case "GLU"
                            If Val(strValues) <= 60 Then
                                strValues = "�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 60 And Val(strValues) <= 100 Then
                                strValues = "1�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 100 And Val(strValues) <= 250 Then
                                strValues = "2�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 250 And Val(strValues) <= 500 Then
                                strValues = "3�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 500 Then
                                strValues = ">=3�� " & Split(strValues, " ")(0)
                            End If
                        Case "PRO"
                            If Val(strValues) <= 15 Then
                                strValues = "�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 15 And Val(strValues) <= 30 Then
                                strValues = "1�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 30 And Val(strValues) <= 100 Then
                                strValues = "2�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 100 And Val(strValues) <= 300 Then
                                strValues = "3�� " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 300 Then
                                strValues = ">=3�� " & Split(strValues, " ")(0)
                            End If
                        Case Else
                            strValues = Split(strValues, " ")(0)
                        End Select
                    End If
                    strItem = strItem & "|" & str��Ŀ & "|" & strValues
                    'End If)
                Next
                If UBound(varLine) > iRow + 14 Then
                    For iReturn = iRow + 15 To UBound(varLine)
                        If iReturn < UBound(varLine) Then
                            strReserved = strReserved & varLine(iReturn) & vbCrLf
                        Else
                            If varLine(iReturn) <> "" Then
                                strReserved = strReserved & varLine(iReturn)
                            End If
                        End If
                    Next
                End If
                Exit For
            Else
                strReserved = strSource
                Exit Sub
            End If
            
        End If
    Next
    
    If Val(str�걾��) > 0 And strItem <> "" Then
        strResult = str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    'strCmd = Chr(6)
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





