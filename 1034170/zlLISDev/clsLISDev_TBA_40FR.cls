VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_TBA_40FR"
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
    
    Dim iStrPoint As Integer, iStrPointBK As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    
    Dim strChenal As String, dblValue As Double, iPos As Long
    
    Dim dbl�ܵ��� As Double, dbl�׵��� As Double, dbl���ܶ�֬���� As Double, dbl���ܶ�֬���� As Double
    Dim dbl��֬����A As Double, dbl��֬����B As Double, dbl�ܵ����� As Double, dblֱ�ӵ����� As Double
    Dim blnEnd As Boolean, strLine As String, str���� As String, str�걾�� As String
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    If strSource Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            
            strLine = GetStr_Section(strReserved, Chr(2), Chr(3))
            strReserved = Mid(strReserved, InStr(strReserved, Chr(3)) + 1)
    
            iStrPoint = 1: blnReadItem = False: str���� = "": str�걾�� = ""
            Do While iStrPoint <= Len(strLine)
                If Not blnReadItem Then '����һ�����ݵĿ�ʼλ��
                    iStrPoint = InStr(iStrPoint, strLine, "")
                    '�ж�һ�����Ƿ�ʼ
                    If iStrPoint > 0 Then
                        '��¼һ���������
                        iStrPoint = iStrPoint + 1
                        iCurrGrpEntry = iStrPoint - 1
                        '�����Ƿ�����
                        If InStr(iStrPoint, strLine, "") > 0 Then
                            '����ʱ��
                            'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                            iStrPoint = iStrPoint + 1
                            strData = Format(Year(Date) & "-" & Format(ReadSubString(strLine, iStrPoint, 4), "00-00"), "yyyy-mm-dd hh:mm:ss")
                            str���� = strData
                            '������ţ�����Ϊ����
                            iStrPoint = iStrPoint + 9
                            
                            str�걾�� = Val(ReadSubString(strLine, iStrPoint, 5))
                            strData = strData & "|" & str�걾��
                            
                            '������
                            strData = strData & "| " '��������֧��
                            '�걾
                            strData = strData & "|ѪҺ" '��������֧��
                            '�ʿ�Ʒ
                            strData = strData & "|0"
                            
                            '׼������
                            iStrPoint = iStrPoint + 24
                            blnReadItem = True
                        
                            dbl�ܵ��� = 0
                            dbl�׵��� = 0
                            dbl���ܶ�֬���� = 0
                            dbl���ܶ�֬���� = 0
                            dbl��֬����A = 0
                            dbl��֬����B = 0
                            dbl�ܵ����� = 0
                            dblֱ�ӵ����� = 0
                            
                        End If
                    Else
                        Exit Do
                    End If
                Else
                    strBuffer = ReadSubString(strLine, iStrPoint, 1)
                    If strBuffer = "" Then '�������
                        If dbl�ܵ��� <> 0 And dbl�׵��� <> 0 Then
                            '�򵰰� �ܵ��� - �׵���
                            strData = strData & "|50"
                            strData = strData & "|" & Format(Round((dbl�ܵ��� - dbl�׵���), 2), "#0.00")
                            'A/G �׵���/�򵰰�
                            strData = strData & "|51"
                            strData = strData & "|" & Format(Round((dbl�׵��� / IIf(dbl�ܵ��� - dbl�׵��� = 0, 1, dbl�ܵ��� - dbl�׵���)), 2), "#0.00")
                        End If
                        '��ӵ����� �� - ֱ��
                        If dbl�ܵ����� <> 0 And dblֱ�ӵ����� <> 0 Then
                            strData = strData & "|52"
                            strData = strData & "|" & Format(Round((dbl�ܵ����� - dblֱ�ӵ�����), 2), "#0.00")
                        End If
                        '��������Ӳ��ָ�� ���ܶ�֬���� / ���ܶ�֬����
                        If dbl���ܶ�֬���� <> 0 And dbl���ܶ�֬���� <> 0 Then
                            strData = strData & "|53"
                            strData = strData & "|" & Format(Round((dbl���ܶ�֬���� / IIf(dbl���ܶ�֬���� = 0, 1, dbl���ܶ�֬����)), 2), "#0.00")
                        End If
                        '��֬����A/B ��֬����A-1/��֬����B
                        If dbl��֬����B <> 0 And dbl��֬����A <> 0 Then
                            strData = strData & "|54"
                            strData = strData & "|" & Format(Round((dbl��֬����A / IIf(dbl��֬����B = 0, 1, dbl��֬����B)), 2), "#0.00")
                        End If
                        
                        'strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                        
                        blnReadItem = False
                        '�������ѽ���
                        iCurrGrpEntry = 0
                    Else
                        '������Ŀ
                        iStrPoint = iStrPoint - 1
                        strChenal = ReadSubString(strLine, iStrPoint, 2)
                        strData = strData & "|" & strChenal
                        dblValue = Round(Val(ReadSubString(strLine, iStrPoint, 7)), 2)
                        strData = strData & "|" & Format(dblValue, "#0.00")
                        Select Case strChenal
                            Case "01"
                                dbl�ܵ��� = dblValue
                            Case "02"
                                dbl�׵��� = dblValue
                            Case "16"
                                dbl���ܶ�֬���� = dblValue
                            Case "15"
                                dbl���ܶ�֬���� = dblValue
                            Case "20"
                                dbl��֬����A = dblValue
                            Case "21"
                                dbl��֬����B = dblValue
                            Case "03"
                                dbl�ܵ����� = dblValue
                            Case "04"
                                dblֱ�ӵ����� = dblValue
                        End Select
                    End If
                End If
            Loop
                    
'            '����δ������ԭʼ��
'            If iCurrGrpEntry > 0 Then
'                strReserved = Mid(strSource, iCurrGrpEntry)
'            Else
'                '���豸���͡������յ�������
'                strCmd = ""
'            End If
            'strResult = strResult & "||" & strData
            If IsDate(str����) And Val(str�걾��) > 0 Then
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
            End If
            If Not strReserved Like "*" & Chr(&H2) & "*" & Chr(&H3) & "*" Then
                blnEnd = True
            End If
        Loop
    Else
        strReserved = strSource
        'strCmd = Chr(&H6&)
    End If
    
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
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