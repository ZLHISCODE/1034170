VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BC_5500"
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
    Dim i As Integer
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
    '06/ 8/1117:50           97ZRSZ   Q  8.3 3.66  1240.358 97.8 33.9  346   890.2090.0550.736  1.7  0.5  6.10.154 55.6---.----.--.--- 51.2161.7               +                                 MPMPMP
    '[kx21]D1U0610125000000000099000000S010400522001240038700741002380032000298003690005400577000380000600060004150010900089001700

    Dim iStrPoint As Integer, intSpot As Integer
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim iPos As Long, strTmp As String, dblValue As Double, strItemValue As String
    Dim strbarcode As String
    Dim varGroup As Variant
    Dim varItem As Variant
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    If strSource Like "*CTR*TestTime*AbnormalFlag*" Then
    
        intSpot = InStr(strSource, "")
        
        iStrPoint = 1
        Do While iStrPoint <= Len(strSource)
            iStrPoint = InStr(iStrPoint, strSource, "CTR")
            If iStrPoint = 0 Then
                iStrPoint = 1
                iStrPoint = InStr(iStrPoint, strSource, "QCR")
            End If
            If iStrPoint = 0 Then
                iStrPoint = 1
                iStrPoint = InStr(iStrPoint, strSource, "QCC")
            End If
            '�ж�һ�����Ƿ�ʼ
            If iStrPoint > 0 Then
                '��¼һ���������
                iCurrGrpEntry = iStrPoint
                
                iPos = InStr(iStrPoint, strSource, "CTR")
                If iPos = 0 Then
                    iPos = InStr(iStrPoint, strSource, "QCR")
                End If
                If iPos = 0 Then
                    iPos = InStr(iStrPoint, strSource, "QCC")
                End If
                If iPos = 0 Then Exit Do
                            
                'û�н���ʱ�˳�
                If InStr(strSource, "AbnormalFlag") = 0 Then Exit Do
                
                If InStr(iPos, strSource, "CTR") > 0 Or InStr(iPos, strSource, "QCR") > 0 Or InStr(iPos, strSource, "QCC") > 0 Then
                    
                    iStrPoint = iStrPoint + 5
                    
                    varGroup = Split(Mid(strSource, iStrPoint, InStr(strSource, "AbnormalFlag") + 14), "")
                    
                    'û���յ�����������
                    If UBound(varGroup) < 44 Then Exit Sub
                    
                    varItem = Split(Mid(varGroup(44), InStr(varGroup(44), "SampleID") + 8), "")
                    
                    '����ʱ��
                    strData = Mid(varGroup(45), InStr(varGroup(45), "") + 1)
                    '�걾��
                    strData = strData & "|" & Mid(varItem(0), InStr(varItem(0), "") + 1)
                    '������
                    strData = strData & "|" & Mid(varItem(3), InStr(varItem(3), "") + 1)
                    '�걾(������֧��)
                    strData = strData & "|Ѫ��"
                    '�Ƿ��ʿ�
                    If InStr(iPos, strSource, "QCC") > 0 Then
                        strData = strData & "|1"
                    Else
                        strData = strData & "|0"
                    End If
                    
                    For i = 0 To 35
                        varItem = Split(varGroup(i), "")
                        
                        If Mid(varItem(0), InStr(varGroup(i), "Val") + 4) <> "***.*" Then
                            strData = strData & "|" & Mid(varItem(0), 1, InStr(varItem(0), Chr(&HC&)) - 1) & "|" & Mid(varItem(0), InStr(varGroup(i), "Val") + 4)
                        End If
                    Next
                    
                    Call ReadSubString(strSource, iStrPoint, InStr(strSource, "AbnormalFlag") - 14)
                    
                    '�������ѽ���
                    iCurrGrpEntry = 0
                    If Left(strResult, 1) = "|" Then strResult = Mid(strResult, 2)
                    strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                Else
                    Exit Do
                End If
              
                'ָ�붨λ����һ������֮��
            Else
                Exit Do
            End If
        Loop
        '����δ������ԭʼ��
        If iCurrGrpEntry > 0 Then
            strReserved = Mid(strSource, iCurrGrpEntry)
        End If
    Else
        strReserved = strSource
    End If
    
    '���豸���͡������յ�������
    strCmd = Chr(&H6&)
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

