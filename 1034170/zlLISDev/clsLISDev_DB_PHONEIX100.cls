VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_DB_PHONEIX100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  ΢���� DB PHONEIX100
'ҽԺ:  �ɶ�����
'����:  ����
'ʱ��:  2008/1/29

Dim strTmp  As String
Private mstrCode As String

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
    '   ΢��������ʽ ͨ����|�������^��ʽ^���ֵ
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
    Dim lngLoop As Long, intloop As Integer
    Dim strDate As String
    Dim i As Integer
    Dim strTmp As String
    Dim strSampleNO As String           '�걾��
    Dim strBarcode As String            '�����
    Dim strImageData As String          'ͼ����Ϣ
    Dim strLine As String               '��ǰҪ�������ִ�
    Dim lngBegin As Long                '��ʼλ��
    Dim lngEnd As String                '����λ��
    
    On Error GoTo errH
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    'û���ҵ���ʼλ��ʱ�˳�
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
    End If
    
    If InStr(strSource, Chr(&HA)) > 0 Then
        strCmd = Chr(&H6)
    End If
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
     
    'û���ҵ�����λ��ʱ�˳�
    lngBegin = InStr(mstrCode, Chr(&H5))
    If lngBegin > 0 Then
        lngEnd = InStr(lngBegin, mstrCode, Chr(&H4))
    Else
        lngEnd = InStr(mstrCode, Chr(&H4))
    End If
    
    Do While lngBegin > 0 Or lngEnd > 0
    
        'ȡ��ǰһ�εĽ���������ķ���
        strLine = Mid(mstrCode, lngBegin, lngEnd - lngBegin + 1)
        mstrCode = Mid(mstrCode, lngEnd + 1)
    
        'ȥ������,�����ַ�
        strLine = DropEccCode(strLine)
        
        '��ʼ����
        aData = Split(strLine, Chr(&HD))
        aItem = Split(aData(0), "|")
        '����
        strData = Mid(aItem(13), 1, 4) & "-" & Mid(aItem(13), 5, 2) & "-" & Mid(aItem(13), 7, 2) & _
                  " " & Mid(aItem(13), 9, 2) & ":" & Mid(aItem(13), 11, 2) & ":" & Mid(aItem(13), 13, 2)
        '�걾��^�걾
        aItem = Split(aData(2), "|")
        aItem = Split(aItem(2), "^")
        strData = strData & "|" & Val(aItem(0)) & "| |" & aItem(2) & "|0"
        
        '��ʼ�������
        For intloop = 0 To UBound(aData)
            If UBound(Split(aData(intloop), "|")) >= 1 Then
                If Split(aData(intloop), "|")(0) = "R" And Val(Split(aData(intloop), "|")(1)) > 1 Then
                    aItem = Split(aData(intloop), "|")
                    If UBound(aItem) < 4 Then
                        aItem = Split(Replace(aData(intloop) & aData(intloop + 1), vbLf, ""), "|")
                    End If
                    If UBound(aItem) >= 3 Then
                        If UBound(Split(aItem(3), "^")) >= 2 Then
                        
                            strData = strData & "|" & Replace(Replace(Split(aItem(2), "^")(5), Chr(&H17), ""), Chr(&H2), "") & _
                                        "|" & Split(aItem(3), "^")(2)
                            If InStr(Split(aItem(2), "^")(3), "MIC") > 0 Then
                                strData = strData & "^1"
                            ElseIf InStr(Split(aItem(2), "^")(3), "DISK") > 0 Then
                                strData = strData & "^2"
                            Else
                                strData = strData & "^3"
                            End If
                            strData = strData & "^" & Split(aItem(3), "^")(1)
                        End If
                    End If
                End If
            End If
        Next
        
        strResult = strResult & "||" & strData
        lngBegin = InStr(mstrCode, Chr(&H5))
        If lngBegin > 0 Then
            lngEnd = InStr(lngBegin, mstrCode, Chr(&H4))
        Else
            lngEnd = InStr(mstrCode, Chr(&H4))
        End If

    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errH:
    WriteLog "err", Err.Description, strResult
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
    
End Function




Private Function DropEccCode(ByVal strCode As String) As String
    'ȥ�������У������Ϣ
    Dim lngS As Long, lngE As Long
    Dim strIn As String, strReturn As String
    strIn = strCode
    strReturn = ""
    Do While Len(strIn) > 0
        lngS = InStr(strIn, Chr(&H2))
        If lngS > 0 Then
            strReturn = strReturn & Mid(strIn, 1, lngS - 1)
            strIn = Mid(strIn, lngS + 2)
        Else
            strReturn = strReturn & strIn
            strIn = ""
        End If
    Loop
    
    strIn = strReturn
    strReturn = ""
    
    Do While Len(strIn) > 0
        lngS = InStr(strIn, Chr(&H17))
        If lngS > 0 Then
            strReturn = strReturn & Mid(strIn, 1, lngS - 1)
            strIn = Mid(strIn, lngS + 5)
        Else
            strReturn = strReturn & strIn
            strIn = ""
        End If
    Loop
    DropEccCode = strReturn
End Function




