VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AU2700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'��Ժ ˫��
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
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean
    Dim intQC As Integer       '�ʿ�λ��
    Dim intDataType As Integer '�������� 0�������1������
    Dim intEnd As Integer      '���λ��
    Dim strResultTmp As String    '��ʱ������
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    
    Dim intE As Integer, intP As Integer, intQ As Integer
    Dim intEmergency As Integer
    
    intE = Val(ReadIni("AU2700", "E", App.Path & "\AU2700.ini"))
    If intE = 0 Then intE = 1000:   Call WriteIni("AU2700", "E", 8000, App.Path & "\AU2700.ini")
    
    intQ = Val(ReadIni("AU2700", "Q", App.Path & "\AU2700.ini"))
    If intQ = 0 Then intE = 2900:   Call WriteIni("AU2700", "Q", 6000, App.Path & "\AU2700.ini")
    
    intP = Val(ReadIni("AU2700", "P", App.Path & "\AU2700.ini"))
    If intP = 0 Then intE = 2000:   Call WriteIni("AU2700", "P", 3000, App.Path & "\AU2700.ini")
    
    intEmergency = Val(ReadIni("AU2700", "�Ƿ�ʹ�ü���", App.Path & "\AU2700.ini"))
    If intEmergency = 0 Then Call WriteIni("AU2700", "�Ƿ�ʹ�ü���", CStr(intEmergency), App.Path & "\AU2700.ini")
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        If Not blnReadItem Then '����һ�����ݵĿ�ʼλ��
            strSource = Replace(strSource, "DB", "")
            strSource = Replace(strSource, "DE", "")
            strSource = Replace(strSource, "D01B", "")
            strSource = Replace(strSource, "D01E", "")
            strSource = Replace(strSource, "RB", "")
            strSource = Replace(strSource, "RE", "")
            
            intDataType = -1
            iPos = InStr(iStrPoint, strSource, "R")
            If iPos > 0 Then
                '��������
                intDataType = 1
            Else
                '�������
                iPos = InStr(iStrPoint, strSource, "D")
                If iPos > 0 Then intDataType = 0
            End If
            Select Case intDataType
                Case 0 '���
                    '��¼һ���������
                    iStrPoint = iPos + 1
                    iCurrGrpEntry = iStrPoint - 1
                    '�����Ƿ�����
                    iPos = InStr(iStrPoint, strSource, "")
                    If iPos > 0 Then
                        'ȥ�����δ�����������Ϣ
                        Do While True
                            iPos = InStr(strSource, "")
                            If iPos = 0 Then Exit Do
                            strSource = Mid(strSource, 1, iPos - 1) & Mid(strSource, iPos + 2 + 41)
                        Loop
                
                        '����ʱ��
                        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                        strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                                
                        '������ţ�����Ϊ����
                        iStrPoint = iStrPoint + 9
                        strBuffer = ReadSubString(strSource, iStrPoint, 4)
                        If Not IsNumeric(Mid(strBuffer, 1, 1)) Then
                            Select Case Mid(strBuffer, 1, 1)
                                Case "E"
                                    strData = strData & "|" & intE + Val(Mid(strBuffer, 2)) & IIf(intEmergency = 0, "^1", "^0")
                                Case "P"
                                    strData = strData & "|" & intP + Val(Mid(strBuffer, 2)) & IIf(intEmergency = 0, "^1", "^0")
                                Case "Q"
                                    blnQc = True
                                    strData = strData & "|" & intQ + Val(Mid(strBuffer, 2)) * 10 + _
                                              Val(Mid(strSource, iStrPoint, 24)) & IIf(intEmergency = 0, "^1", "^0")
                            End Select
                        Else
                            strData = strData & "|" & Val(strBuffer) & "^0"
                        End If
                        'SampleID������,20λ
                        strData = strData & "^" & ReadSubString(strSource, iStrPoint, 20)
                                
                        '������
                        strData = strData & "| " '��������֧��
                        '�걾
                        strData = strData & "|Ѫ��" '��������֧��
                        '�ʿ�Ʒ
                        If blnQc = True Then
                            strData = strData & IIf(blnQc, "|1", "|0")
                            intEnd = InStr(iStrPoint, strSource, "E") + 2
                            strBuffer = Val(ReadSubString(strSource, iStrPoint, intEnd - iStrPoint))
                            Select Case Val(strBuffer)
                                Case 1
                                    strData = strData & "^" & 402
                                Case 2
                                    strData = strData & "^" & 403
                                Case 3
                                    strData = strData & "^" & 404
                                Case Else
                                    strData = strData & "^" & Val(strBuffer)
                            End Select
                            
                        Else
                            intEnd = InStr(iStrPoint, strSource, "E")
                            strData = strData & "|0"
                            iStrPoint = iStrPoint + (intEnd - iStrPoint)
                        End If
                        'strData = strData & "|0" 'IIf(Mid(strSource, iStrPoint, 1) = "E", "|1", "|0")
                        
                        '׼������
                        iStrPoint = iStrPoint + 1
                        blnReadItem = True
                    Else
                        Exit Do
                    End If
                Case 1 '����
                    '��¼һ���������
                    iStrPoint = iPos + 1
                    iCurrGrpEntry = iStrPoint - 1
                    '�����Ƿ�����
                    iPos = InStr(iStrPoint, strSource, "")
                    If iPos > 0 Then
                        'ȥ�����δ�����������Ϣ
                        Do While True
                            iPos = InStr(strSource, "")
                            If iPos = 0 Then Exit Do
                            
                            strSource = Mid(strSource, 1, iPos - 1) & Mid(strSource, iPos + 2 + 41)
                        Loop
                
                        '����ʱ��
                        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                        strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                                
                        '������ţ�����Ϊ����
                        iStrPoint = iStrPoint + 9
                        strBuffer = ReadSubString(strSource, iStrPoint, 4)
                        If Not IsNumeric(Mid(strBuffer, 1, 1)) Then
                            strData = strData & "|" & Val(Mid(strBuffer, 2)) & "^1"
                        Else
                            strData = strData & "|" & Val(strBuffer) & "^0"
                        End If
                        'SampleID������,20λ
                        strData = strData & "^" & ReadSubString(strSource, iStrPoint, 20)
                                
                        '������
                        strData = strData & "| " '��������֧��
                        '�걾
                        strData = strData & "|Ѫ��" '��������֧��
                        '�ʿ�Ʒ
                        strData = strData & "|0" 'IIf(Mid(strSource, iStrPoint, 1) = "E", "|1", "|0")
                        
                        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                        '�������ѽ���
                        iCurrGrpEntry = 0
                        If InStr(iStrPoint, strSource, "R") = 0 Then Exit Do
                    Else
                        Exit Do
                    End If
                Case Else
                    Exit Do
            End Select
        Else
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
            If strBuffer = "" Then '�������
'                '������򵰰�(3-4)��ͨ���룺50
'                iPos = InStr(28, strData, "|03|")
'                sng1 = IIf(iPos = 0, 0, Val(Mid(strData, iPos + 4, 8)))
'                iPos = InStr(28, strData, "|04|")
'                sng2 = IIf(iPos = 0, 0, Val(Mid(strData, iPos + 4, 8)))
'                strData = strData & "|50|" & Mid(CStr(sng1 - sng2), 1, 8)
'                '�ǻ�Ѫ�쵰��()
                If UBound(Split(strData, "|")) > 5 Then
                    strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                End If
                blnReadItem = False
                '�������ѽ���
                iCurrGrpEntry = 0
            Else
                '������Ŀ
                iStrPoint = iStrPoint - 1
                strData = strData & "|" & ReadSubString(strSource, iStrPoint, 2)
                strResultTmp = ReadSubString(strSource, iStrPoint, 11)
                strResultTmp = Replace(strResultTmp, "*", "")
                strResultTmp = Replace(strResultTmp, "r", "")
                
                If InStr(strResultTmp, "!") > 0 Then
                    strData = strData & "|" & Val(Replace(strResultTmp, "%", "")) & "^!^^"
                Else
                    strData = strData & "|" & Val(Replace(Replace(strResultTmp, "%", ""), "?", ""))
                End If
                
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
        Select Case intDataType
            Case 0 '���ؽ��
                strCmd = "0|" & Chr(&H6&)
            Case 1 '������������
                strCmd = "1|"
            Case Else
                strCmd = ""
        End Select
    End If
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
    Dim blnEnd As Boolean
    Dim strBuffer As String, strSampleID As String
    Dim i As Long, aItem() As String
    Dim strRackNO As String, iPos As Long
    
    blnSuccess = True
    Select Case iSendStep
'        Case 0 '��ʼ����
'            SendSample = Chr(6)
'            iSendStep = 1
        Case Else '��������
            If strResponse Like "*R*" Or Len(Trim(strResponse)) = 0 Then '��������
                If strResponse Like "**" Or Len(Trim(strResponse)) = 0 Then
                    '�������
                    '�����豸��
                    aItem = Split(strSample, "|")
                    '�걾��SampleID
                    strBuffer = aItem(1)
                    strSampleID = aItem(5)
                    '��ȡ�����ļܺźͱ���
                    strRackNO = ""
                    If Len(Trim(strResponse)) > 0 Then
                        strResponse = Replace(strResponse, "RB", "")
                        iPos = InStr(strResponse, "R")
                        strRackNO = Mid(strResponse, iPos + 3, 11)
                    End If
                    If Len(Trim(strSampleID)) > 0 Then
                        'SampleID
                        SendSample = "S " & Right(Space(11) & strRackNO, 11) & Right(Space(20) & strSampleID, 20) & Space(4) & "E"
                    Else
                        SendSample = "S        " & Right(String(4, "0") & strBuffer, 4) & Space(20) & Space(4) & "E"
                    End If
                    For i = 10 To UBound(aItem)
                        SendSample = SendSample & aItem(i)
                    Next
                    SendSample = SendSample & "SE"
                
                    iSendStep = 0
                Else
                    SendSample = Chr(6)
                    iSendStep = iSendStep + 1
                End If
            Else
                SendSample = Chr(6)
                iSendStep = 0
                blnSuccess = False '����ʧ��
            End If
    End Select
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



