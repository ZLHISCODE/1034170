VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_diff5"
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
    
    Dim iStrPoint As Integer, iStrPointBK As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim aData() As String, iDataPoint As Integer, aItem() As String
    
    Dim strChenal As String, dblValue As Double, iPos As Long, iStartPoint As Integer
    Dim str�걾��Ϣ As String
    Dim dbl�ܵ��� As Double, dbl�׵��� As Double, dbl���ܶ�֬���� As Double, dbl���ܶ�֬���� As Double
    Dim dbl��֬����A As Double, dbl��֬����B As Double, dbl�ܵ����� As Double, dblֱ�ӵ����� As Double
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        If Not blnReadItem Then '����һ�����ݵĿ�ʼλ��
            iStrPoint = InStr(iStrPoint, UCase(strSource), "RESULT")
            '�ж�һ�����Ƿ�ʼ
            If iStrPoint > 0 Then
                '��¼һ���������
                iCurrGrpEntry = iStrPoint
                iStrPoint = iStrPoint + 6
                '�����Ƿ�����
                If InStr(iStrPoint, UCase(strSource), "ACT5DIFF") > 0 Then
                    aData = Split(Mid(strSource, iStrPoint), vbCr)
                    '�����Ƿ�����
                    If UBound(aData) > 7 Then
                        '����ʱ��
                        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                        strData = Format(Split(aData(47), " ")(1) & " " & Time, "yyyy-mm-dd hh:mm:ss")
                                
                        '������ţ�����Ϊ����
                        strData = strData & "|" & Val(Split(aData(50), " ")(1))
                                
                        '������
                        strData = strData & "| " '��������֧��
                        '�걾
                        strData = strData & "|ѪҺ" '��������֧��
                        '�ʿ�Ʒ
                        strData = strData & "|0"
                        str�걾��Ϣ = strData
                        '׼������
                        iPos = InStr(iStrPoint, strSource, "!")
                        If iPos > 0 Then
                            aData = Split(Mid(strSource, iPos), vbCr)
                            iDataPoint = 0: iStartPoint = 0
                        Else
                            iDataPoint = 8: iStartPoint = 8
                        End If
                        blnReadItem = True
                    End If
                End If
            Else
                Exit Do
            End If
        Else
            If InStr(UCase(aData(iDataPoint)), "ACT5DIFF") > 0 Then   '�������
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                
                blnReadItem = False
                '�������ѽ���
                iCurrGrpEntry = 0
            Else
                '������Ŀ
                strChenal = Format(iDataPoint - iStartPoint + 1, "00")
                If strChenal > "26" Then 'ͼ������ ALT�ֲ�|0| | | |001001231231223
'                   Select Case Mid(aData(iDataPoint), 1, 1)
'                        Case 19
'                            strData = strData & Chr(13) & Chr(10) & Trim(Mid(aData(iDataPoint), 1, 2)) & "|0| | | |" & to_Hex(Mid(aData(iDataPoint), 3))
'                        Case 20, 21
'                            strData = strData & "||" & str�걾��Ϣ & Chr(13) & Chr(10) & Trim(Mid(aData(iDataPoint), 1, 2)) & "|0| | | |" & to_Hex(Mid(aData(iDataPoint), 3))
'                        End Select
                       
                        iCurrGrpEntry = 0
                Else
                    strData = strData & "|" & strChenal
                    Select Case strChenal
                        Case 17, 21
                            dblValue = Val(Mid(aData(iDataPoint), 3)) * 10
                        Case 19
                            dblValue = Val(Mid(aData(iDataPoint), 3)) * 1.1
                        Case Else
                            dblValue = Val(Mid(aData(iDataPoint), 3))
                        End Select
                    strData = strData & "|" & dblValue
                End If
                iDataPoint = iDataPoint + 1
            End If
        End If
    Loop
            
    '����δ������ԭʼ��
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    Else
        '���豸���͡������յ�������
        strCmd = ""
    End If
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
    Dim lngAsc As Long, I As Integer
    
    For I = 1 To Len(strSource)
        lngAsc = lngAsc + Asc(Mid(strSource, I, 1))
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

Private Function to_Hex(str��¼ As String) As String
Dim I  As Integer
    to_Hex = ""
    For I = 1 To Len(str��¼)
            to_Hex = to_Hex & Hex(Asc(Mid(str��¼, I, 1)))
    Next
End Function











