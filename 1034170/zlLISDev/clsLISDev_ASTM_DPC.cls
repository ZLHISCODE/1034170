VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ASTM_DPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private strDeviceString '������Ϣ���豸��
Private strLastString As String '�ϴη��͵��豸��

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
    
    Dim iStrPoint As Long, iStrPointBK As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean
    Dim aItemData() As String, strFldDelim As String, strRepeatDelim As String, strCompDelim As String
    Dim iMode As Integer '1:Current Mode ��2:New Mode
    Dim iResultType As Integer '������ͣ�1����������2����������
    Dim i As Long, strChannel As String
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    iResultType = 1 '��ʼ�Ƿ��ؽ��
    
    iStrPoint = 1: blnReadItem = False
    'ȡ�������
    strSource = Replace(strSource, Chr(5), "")
    'ȡ��������
    strSource = Replace(strSource, Chr(4), "")
    'ȡ����ʼ��
    strSource = Replace(strSource, Chr(2), "")
    'ȡ���س����з�
    strSource = Replace(strSource, vbCrLf, "")
    'ȡ�����з�
    strSource = Replace(strSource, Chr(10), "")
    'ȡ���Ʊ���
    strSource = Replace(strSource, Chr(9), "")
    'ȡ��ͨ�����ӷ�
    i = 1
    Do While True
        i = InStr(i, strSource, Chr(3))
        If i = 0 Then Exit Do
        If i + 2 <= Len(strSource) Then
            strSource = Mid(strSource, 1, i - 1) & Mid(strSource, i + 3)
        Else
            Exit Do
        End If
    Loop
    
    Do While iStrPoint <= Len(strSource)
        '������Ϣ�Ŀ�ʼλ�ã���H��ʼ
        iPos = iStrPoint
        strBuffer = LTrim(ReadLine(strSource, iPos))
        If Mid(strBuffer, 1, 2) Like "*H*" Then
            '��¼һ���������
            iCurrGrpEntry = iStrPoint
            
            '��¼�ָ���
            If Mid(strBuffer, 1, 2) Like "*H" Then
                strFldDelim = Mid(strBuffer, 3, 1) '�ֶηָ���
                strRepeatDelim = Mid(strBuffer, 4, 1) '�ظ����ݷָ���
                strCompDelim = Mid(strBuffer, 5, 1) '������ݷָ���
            Else
                strFldDelim = Mid(strBuffer, 2, 1) '�ֶηָ���
                strRepeatDelim = Mid(strBuffer, 3, 1) '�ظ����ݷָ���
                strCompDelim = Mid(strBuffer, 4, 1) '������ݷָ���
            End If
            
            '�����Ƿ�����
            If InStr(iStrPoint, strSource, "L" & strFldDelim & "1") > 0 Then
                '��Ϣ���������ٱ���
                iCurrGrpEntry = 0
                
                '���ұ걾��Ϣ
                iStrPoint = iPos: blnReadItem = False
                aItemData = Split(strBuffer, strFldDelim)
                '����ʱ��
                        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                        strData = aItemData(13)
                        strData = Format(Mid(strData, 1, 4) & "-" & Mid(strData, 5, 2) & "-" & Mid(strData, 7, 2) & " " & Mid(strData, 9, 2) & ":" & Mid(strData, 11, 2) & ":" & Mid(strData, 13, 2), "yyyy-mm-dd hh:mm:ss")
                
                Do While iStrPoint <= Len(strSource)
                    If Not blnReadItem Then '��ǰ�Ƿ����ڶ�ȡ������Ŀ
                        Do While iStrPoint <= Len(strSource)
                            strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                            If Mid(strBuffer, 2, 1) = "P" Then Exit Do
                        Loop
                        If iStrPoint > Len(strSource) Then
                            '��ǰ��Ϣ��Ч�����ٴ�����
                            Exit Do
                        End If
                        
                        aItemData = Split(strBuffer, strFldDelim)
                        
                        '������ţ�����Ϊ����
                        strData = strData & "|" & Val(aItemData(2))
                                
                        '������
                        strData = strData & "| " '��������֧��
                        '�걾
                        strData = strData & "|ѪҺ"
                        '�ʿ�Ʒ
                        strData = strData & "|0"
                        
                        '׼������
                        blnReadItem = True
                    Else
                        Do While iStrPoint <= Len(strSource)
                            iPos = iStrPoint '������ǰλ��
                            strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                            If Mid(strBuffer, 2, 1) = "R" Then
                            '������
                                aItemData = Split(strBuffer, strFldDelim)
                                'strChannel = Split(aItemData(2), strCompDelim)(3) & Split(aItemData(2), strCompDelim)(4)
'                                strData = strData & "|" & Format(Val(aItemData(3)), "0.00")
                                strData = strData & "|" & IIf(IsNumeric(aItemData(3)), Format(Val(aItemData(3)), "0.00"), aItemData(3))
                            Else
                                If Mid(strBuffer, 2, 1) <> "O" Then
                                    blnReadItem = False
                                    Exit Do
                                Else
                                    aItemData = Split(strBuffer, strFldDelim)
                                    strChannel = aItemData(1) 'Replace(strChannel, "/", ""): strChannel = Replace(strChannel, "-", "")
                                    strData = strData & "|" & strChannel
                                
                                End If
                            End If
                        Loop
                        'һ���걾����
                        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                        
                        iStrPoint = iPos
                        If InStr("PO", Mid(strBuffer, 2, 1)) = 0 Then
                        '������һ���걾
                            Exit Do
                        End If
                    End If
                Loop
            Else
                Exit Do
            End If
        Else
            iStrPoint = iPos
        End If
    Loop
            
    '����δ������ԭʼ��
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    End If
    '���豸���͡������յ�������
    strCmd = Chr(&H6&)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   �걾֮����||�ָ�
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    blnSuccess = True
    iSendStep = 0
    SendSample = ""
End Function

Private Function GetCheckCode(ByVal strSource As String) As String
'��ȡУ����
    Dim lngAsc As Long, i As Integer
    
    For i = 1 To Len(strSource)
        lngAsc = lngAsc + Asc(Mid(strSource, i, 1))
    Next
    GetCheckCode = CStr(Hex(lngAsc Mod 256))
End Function

Private Function ReadSubString(ByVal strSource As String, iPoint As Long, Optional ByVal Length As Integer = 0) As String
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

Private Function ReadLine(ByVal strSource, iPoint As Long) As String
    '��ָ��λ�ÿ�ʼ��ȡһ�У�ͬʱָ�붨λ����һ������֮��
    Dim iPos As Long
    
    iPos = iPoint
    Do While iPos <= Len(strSource)
        If Mid(strSource, iPos, 1) = Chr(13) Then
            Exit Do
        End If
        iPos = iPos + 1
    Loop
    If iPos > Len(strSource) Then
        ReadLine = Mid(strSource, iPoint)
        iPoint = iPos
    Else
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 1
    End If
End Function
