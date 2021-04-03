VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ROCHE_UR"
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
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim blnEndSample As Boolean '�걾����
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean
    Dim aDatas() As String, strItemResult As String, i As Integer, strFlag As String
    Dim lngNo As Long
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        If Not blnReadItem Then '����һ�����ݵĿ�ʼλ��
            iStrPoint = InStr(iStrPoint, strSource, "")
            '�ж�һ�����Ƿ�ʼ
            If iStrPoint > 0 Then
                '��¼һ���������
                iStrPoint = iStrPoint + 1
                iCurrGrpEntry = iStrPoint - 1
                '�����Ƿ�����
                iPos = InStr(iStrPoint, strSource, "")
                If iPos > 0 Then
                    '���ҵ�ǰ�걾��ʼλ��
                    Do While iStrPoint < iPos
                        strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                        If UCase(Mid(strBuffer, 1, 7)) = "SEQ.NO." Then Exit Do
                    Loop
                    If iStrPoint < iPos Then
                        '��ǰ�걾��Ϣ��Ч
                        lngNo = Val(Mid(strBuffer, 8))
                        '����ʱ��
                        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                        strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                        aDatas = Split(Split(strBuffer, " ")(0), ".")
                        strData = Format(aDatas(2) & "-" & aDatas(1) & "-" & aDatas(0) & " " & Time, "yyyy-mm-dd hh:mm:ss")
                                
                        '������ţ�����Ϊ����
                        strData = strData & "|" & lngNo
                                
                        '������
                        strData = strData & "| " '��������֧��
                        '�걾
                        strData = strData & "|��Һ" '��������֧��
                        '�ʿ�Ʒ
                        strData = strData & "|0"
                        
                        '׼������
                        blnReadItem = True
                    End If
                End If
            Else
                Exit Do
            End If
        Else
            iStrPointBK = iStrPoint
            strBuffer = ReadLine(strSource, iStrPointBK)
            
            '�걾�Ƿ����
            iPos = InStr(strBuffer, "")
            blnEndSample = True
            If iPos = 0 Then iPos = Len(strBuffer) + 1: blnEndSample = False
            
            strBuffer = Mid(strBuffer, 1, iPos - 1)
            If Len(strBuffer) > 0 Then '�������
                strData = strData & "|" & Trim(Mid(strBuffer, 2, 4))
                
                strItemResult = Trim(Mid(strBuffer, 7, Len(strBuffer) - 12))
                strFlag = Trim(Mid(strBuffer, 19))
                'ȡ��������λ
                For i = Len(strItemResult) To 1 Step -1
                    If Mid(strItemResult, i, 1) = " " Or IsNumeric(Mid(strItemResult, i, 1)) Then Exit For
                Next
                If i = 0 Then
                    strItemResult = IIf(Left(strBuffer, 1) = "*", IIf(strFlag = "", "+", strFlag), "-") '& strItemResult
                Else
                    strItemResult = Trim(Mid(strItemResult, 1, i)) & IIf(Left(strBuffer, 1) = "*", IIf(strFlag = "", "(+)", "(" & strFlag & ")"), "")
                End If
'                If Not IsNumeric(strItemResult) Then
'                    strItemResult = IIf(UCase(strItemResult) Like "NEG*" Or UCase(strItemResult) Like "NORM*", "-", strItemResult)
'                End If
                strData = strData & "|" & strItemResult
            End If
            
            If blnEndSample Then '�걾�ѽ���
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                
                blnReadItem = False
                '�������ѽ���
                iCurrGrpEntry = 0
            Else
                iStrPoint = iStrPointBK
            End If
        End If
    Loop
            
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
    If iPos > 0 Then
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 2
    Else
        iPos = InStr(iPoint, strSource, Chr(13))
        If iPos = 0 Then iPos = InStr(iPoint, strSource, Chr(10))
        If iPos = 0 Then iPos = Len(strSource) + 1
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 1
    End If
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '��ȡһ��������Ŀ���ƣ������ַ�ָ���Ƶ���Ŀ����֮��
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function
