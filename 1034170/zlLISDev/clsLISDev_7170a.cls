VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_7170a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
    'AnalysestrSource = ":s 1010 1                3 060506103610 6 315.9  8  69.8  9  39.9 10   2.8 11     4 12    94 14   1.2 15   4.0 16   7.0 24    14 "
    Dim iStrPoint As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    Dim strTP As Double, stralb As Double  '����������Ŀ��ԭʼֵ 25-26
    Dim str���� As Double, str���� As Double  '34/39
    Dim strͨ��, str��� As String
    Dim str�򵰰�, strѪ���, str�ߵ��� As Double, str���� As String
    Dim bln���� As Boolean, blnQc As Boolean
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    Dim int��� As Integer
    
    blnQc = False
    iStrPoint = 1
    Do While iStrPoint <= Len(strSource)
        strBuffer = ReadSubString(strSource, iStrPoint, 1)
'        If InStr(strSource, "") <= 0 Then
'            strReserved = strSource
'            Exit Do
'        End If
        '�ж�һ�����Ƿ�ʼ
        If strBuffer = "" Then
            '��¼һ���������
            
            strTP = 0: stralb = 0: str���� = 0: str���� = 0: bln���� = False: str���� = ""
            iCurrGrpEntry = iStrPoint - 1 '63530918
            '�жϳ����뼱��
            iStrPoint = iStrPoint + 1
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
            Select Case UCase(strBuffer)
                Case "N"
                    str���� = ""
                Case "Q"
                    str���� = "^1"
                Case "F"
                    str���� = ""
                    blnQc = True
            End Select
            iStrPoint = iStrPoint + 26
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
                If Len(strSource) - iStrPoint >= 20 Then
                    If Not blnQc Then
                    '����ʱ��
                        iStrPoint = iStrPoint + 1
                        strData = ReadSubString(strSource, iStrPoint, 2)   '��
                        strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '��
                        strData = ReadSubString(strSource, iStrPoint, 2) & "-" & strData '��
                        strData = strData & " " & ReadSubString(strSource, iStrPoint, 2) & ":" & ReadSubString(strSource, iStrPoint, 2)
                        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                        strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                        
                        '������ţ�����Ϊ����
                        iStrPoint = iStrPoint - 36
                    Else
                        strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                        iStrPoint = iStrPoint - 23
                    End If
                    '������
'                    iStrPoint = iStrPoint + 1
                    strSampleID = Trim(ReadSubString(strSource, iStrPoint, 4)) & str����
                    'strData = strData & "|" & IIf(Len(strSampleID) = 0, strTmpID, strSampleID)
                    strData = strData & "|" & strSampleID
                    '������
                    strData = strData & "| " '��������֧��
                    '�걾
                    strData = strData & "|ѪҺ" '��������֧��
                    '�ʿ�Ʒ
                        If blnQc Then
                           strData = strData & "|1"
                           iStrPoint = iStrPoint + 33
                        Else
                           strData = strData & "|0"
                           iStrPoint = iStrPoint + 35
                        End If
                    
                    '������Ŀ
                    int��� = 0
                    Do While ReadSubString(strSource, iStrPoint, 1) <> ""
                        If int��� > 50 Then Exit Do
                        iStrPoint = iStrPoint - 1
                        strͨ�� = Trim(ReadSubString(strSource, iStrPoint, 3))
                        strData = strData & "|" & strͨ��
                        str��� = Val(Trim(ReadSubString(strSource, iStrPoint, 7)))
                        strData = strData & "|" & str���
                        Select Case strͨ��
                               Case 25
                                   strTP = str���
                               Case 26
                                   stralb = str���
                               Case 34
                                   str���� = str���
                               Case 39
                                   str���� = str���
                        End Select
                    int��� = int��� + 1
                    Loop
                        '�������ѽ���
                    If Not blnQc Then
                           str�򵰰� = strTP - stralb
                        If strTP <> 0 And stralb <> 0 Then
                           strѪ��� = Format(stralb / str�򵰰�, "0.00")
                           strData = strData & "|GLO|" & str�򵰰� & "|A/G|" & strѪ���
                        End If
                        If str���� <> 0 And str���� <> 0 Then
                           str�ߵ��� = Format(str���� / str����, "0.00")
                           strData = strData & "|H/L|" & str�ߵ���
                        End If
                    End If
                    iCurrGrpEntry = 0
                    strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                    
                End If
            'ָ�붨λ����һ������֮��
         '   MoveToNextLine strSource, iStrPoint
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

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
        If Mid(strSource, iPoint, 1) = "]" Then
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







