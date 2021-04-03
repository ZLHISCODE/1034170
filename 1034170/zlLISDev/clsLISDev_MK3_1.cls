VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MK3_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const LABITEMS As String = ",pH,PCO2,PO2,pHt,PCO2t,PO2t,ctHb,O2Hb,HHb,COHb,MetHb,SulfHb,pHst,cHCO3st,cHCO3,ctCO2(P),ctCO2(B),BE,BEecf,BB,SO2,PAO2t,AaDO2t,a/AO2t,RI,"
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
    '����ͼ��ķ�ʽ��
    '                   1.ͼ�����ݸ���ָ�����ݺ�ʹ�ûس����з����ָ���
    '                   2.�ж��ͼ������ʱʹ��"^"���ָ�
    '                   3.����ͼ�����ݸ�ʽ: ͼ�񻭷� 0=ֱ��ͼ  1=ɢ��ͼ
    '                     a) ֱ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ);X1;X2;X3;X4;X5...
    '                     b) ɢ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ):
    '                        ��:00000100001000010000100010;00000100001000010000100010;
    '                        ˵��:1.ɢ��ͼ�Ե���ʽ����ÿһ��ʹ�÷ֺ����ָ�.
    '                             2.�ж��ٸ��ֺž��ж�����
    '                             3.ÿһ���ж��ٸ�����ÿһ�еĳ�����ȷ��
    '                             3.��ͼ�ķ����Ǵ����ϱ����»�������65*65��ͼ���Ǵ�65�п�ʼ��(���ϱ߿�ʼ��)
    
    Dim iStrPoint As Integer, iStrPointBK As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim aData() As String, iDataPoint As Integer, aItem() As String
    
    Dim strChenal As String, dblValue As Double, iPos As Long, iStartPoint As Integer
    Dim lngloop As Long, intloop As Integer
    Dim strDate As String
    Dim i As Integer
    Dim strTmp As String
    Dim strSampleNO As String           '�걾��
    Dim strBarcode As String            '�����
    Dim strImageData As String          'ͼ����Ϣ
    Dim strLine As String               '��ǰҪ�������ִ�
    Dim lngBegin As Long                '��ʼλ��
    Dim lngEnd As String                '����λ��
    
    'û����ô��ʱ�˳�
    If UBound(Split(strSource, vbCrLf)) < 8 Then
        strReserved = strSource
        Exit Sub
    End If
    
    '��ʼ����
    strResult = strSource
    For intloop = 5 To 2 Step -1
        strResult = Replace(strResult, Space(intloop), Space(1))
    Next
    strResult = Replace(strResult, "OK" & vbCrLf, "")
    strResult = Replace(strResult, " ", ";")
    strResult = Replace(strResult, "-", ";-")
    strResult = Mid(strResult, 2)
    strResult = Replace(strResult, vbCrLf, "|")
    strResult = Replace(strResult, "|;", "|")
    
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
    
    iPos = InStr(iPoint, strSource, Chr(10))
    If iPos = 0 Then iPos = Len(strSource) + 1
    
    ReadLine = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '��ȡһ��������Ŀ���ƣ������ַ�ָ���Ƶ���Ŀ����֮��
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function



Public Function CmdAnalyse(strCmdType As String) As String
    '����:  �ֽ�����Ϊ����ִ�е�����
    
'    strControl = Me.cbo����.Text & ";" & Me.cbo�ο�����.Text & ";" & Me.cbo���Ƶ��.Text & ";" & Me.txt���ʱ�� & _
'                 ";" & Me.cbo���巽ʽ.Text & ":" & Me.cbo�հ���ʽ.Text
    
    Dim lngTime As Long
    Select Case strCmdType
        Case "Begin"                   '��ʼ
            CmdAnalyse = "R" & vbCrLf
        Case "End"                      '����
            CmdAnalyse = "Q" & vbCrLf
        Case "Revert"                   '�ظ�����
            CmdAnalyse = "OK" & vbCrLf
        Case "ReadData"                  '�������
            CmdAnalyse = "P" & vbCrLf
        Case "450"
            CmdAnalyse = "F2" & vbCrLf  '450��Ƭ
        Case "1(450)"                   '1�Ų�Ƭ
            CmdAnalyse = "F1" & vbCrLf
        Case "2(500)"                   '2�Ų�Ƭ
            CmdAnalyse = "F2" & vbCrLf
        Case "3(750)"                   '3�Ų�Ƭ
            CmdAnalyse = "F3" & vbCrLf
        Case "2(450)"
            CmdAnalyse = "F2" & vbCrLf  '2�벨ƬΪ"450"
        Case "F1"                        '1�Ų�Ƭ
            CmdAnalyse = "F1" & vbCrLf
        Case "F2"                        '2�Ų�Ƭ
            CmdAnalyse = "F2" & vbCrLf
        Case "F3"                        '3�Ų�Ƭ
            CmdAnalyse = "F3" & vbCrLf
        Case "F4"                        '4�Ų�Ƭ
            CmdAnalyse = "F4" & vbCrLf
        Case "F5"                        '5�Ų�Ƭ
            CmdAnalyse = "F5" & vbCrLf
        Case "����"                     '���巽ʽ����
            CmdAnalyse = "E0" & vbCrLf
        Case "��"                     '���巽ʽ��
            CmdAnalyse = "E1" & vbCrLf
        Case "1X"                       '���Ƶ�� 1X
            CmdAnalyse = "X1" & vbCrLf
        Case "2X"
            CmdAnalyse = "X2" & vbCrLf
        Case "3X"
            CmdAnalyse = "X3" & vbCrLf
        Case "X1"                       '���Ƶ�� 1X
            CmdAnalyse = "X1" & vbCrLf
        Case "X2"
            CmdAnalyse = "X2" & vbCrLf
        Case "X3"
            CmdAnalyse = "X3" & vbCrLf
        Case "����"
            CmdAnalyse = "A" & vbCrLf
        Case "�Լ�"
            CmdAnalyse = "B" & vbCrLf
        Case "Z5_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5 & vbCrLf
        Case "Z10_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5 & vbCrLf
        Case "Z15_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5 & vbCrLf
        Case "Z20_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5 & vbCrLf
        Case "Z25_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5 & vbCrLf
        Case "Z30_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5 & vbCrLf
        Case Else
            If Left(strCmdType, 1) = "Z" Then
                '���ʱ��
                lngTime = Val(Mid(strCmdType, 2))
                If lngTime >= 0 And lngTime <= 60 Then
                    CmdAnalyse = "Z" & Format(lngTime, "00") & vbCrLf
                End If
            End If
    End Select
    
End Function






