VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_SF3000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����
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
    
    Dim iStrPoint As Integer, iStrPointBK As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String, strItemValue As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean, lngItemID As Long
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        If Not blnReadItem Then '����һ�����ݵĿ�ʼλ��
            strSource = Replace(strSource, "DB", "")
            strSource = Replace(strSource, "DE", "")
            strSource = Replace(strSource, "D01B", "")
            strSource = Replace(strSource, "D01E", "")
            
            iStrPoint = InStr(iStrPoint, strSource, "D")
            '�ж�һ�����Ƿ�ʼ
            If iStrPoint > 0 Then
                '��¼һ���������
                iStrPoint = iStrPoint + 1
                iCurrGrpEntry = iStrPoint - 1
                '�����Ƿ�����
                iPos = InStr(iStrPoint, strSource, "")
                If iPos > 0 Then
                    '����ʱ��
                    'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                    strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                            
                    '������ţ�����Ϊ����
                    iStrPoint = iStrPoint + 15
                    strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 5))
                            
                    '������
                    strData = strData & "| " '��������֧��
                    '�걾
                    strData = strData & "|ѪҺ" '��������֧��
                    '�ʿ�Ʒ
                    strData = strData & "|0" 'IIf(Mid(strSource, iStrPoint, 1) = "E", "|1", "|0")
                    
                    '׼������
                    iStrPoint = iStrPoint + 33
                    blnReadItem = True
                    lngItemID = 1
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Else
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
            If strBuffer = "" Then '�������
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                
                blnReadItem = False
                '�������ѽ���
                iCurrGrpEntry = 0
            Else
                '������Ŀ
                iStrPoint = iStrPoint - 1
                Select Case lngItemID
                    Case 3, 7, 8
                        strItemValue = ReadSubString(strSource, iStrPoint, 5)
                        If IsNumeric(strItemValue) Then
                            strData = strData & "|" & lngItemID
                            strData = strData & "|" & Val(strItemValue) / 10
                        ElseIf InStr(strItemValue, "*") > 0 Then
                            strData = strData & "|" & lngItemID
                            strData = strData & "|*"
                        End If
                    Case 1, 2, 14, 15, 16, 17, 18
                        strItemValue = ReadSubString(strSource, iStrPoint, _
                            IIf((lngItemID >= 14 And lngItemID <= 18) Or lngItemID = 1, 6, 5))
                        If IsNumeric(strItemValue) Then
                            strData = strData & "|" & lngItemID
                            strData = strData & "|" & Val(strItemValue) / 1000
                        ElseIf InStr(strItemValue, "*") > 0 Then
                            strData = strData & "|" & lngItemID
                            strData = strData & "|*"
                        End If
                    Case 4, 19, 23
                        strItemValue = ReadSubString(strSource, iStrPoint, 5)
                        If IsNumeric(strItemValue) Then
                            strData = strData & "|" & lngItemID
                            strData = strData & "|" & Val(strItemValue) / 10000
                        ElseIf InStr(strItemValue, "*") > 0 Then
                            strData = strData & "|" & lngItemID
                            strData = strData & "|*"
                        End If
                    Case 5, 6, 9, 10, 11, 12, 13, 20, 21, 22
                        strItemValue = ReadSubString(strSource, iStrPoint, 5)
                        If IsNumeric(strItemValue) Then
                            strData = strData & "|" & lngItemID
                            strData = strData & "|" & Val(strItemValue) / 100
                        ElseIf InStr(strItemValue, "*") > 0 Then
                            strData = strData & "|" & lngItemID
                            strData = strData & "|*"
                        End If
                    Case Else
                        iStrPoint = iStrPoint + 1
                End Select
                
                lngItemID = lngItemID + 1
            End If
        End If
    Loop
            
'    If blnReadItem Then
'        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
'        '�������ѽ���
'        iCurrGrpEntry = 0
'    End If
    '����δ������ԭʼ��
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    Else
        '���豸���͡������յ�������
        strCmd = Chr(&H6&)
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