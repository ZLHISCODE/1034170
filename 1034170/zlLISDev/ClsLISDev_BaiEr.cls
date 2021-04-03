VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_BaiEr_500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  �ݶ�500
'ҽԺ:  �ɶ�����
'����:  ����
'ʱ��:  2008/1/16
Private str_Count As String '��¼�յ����ַ���������ʱӦ��

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
    
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    'û���ҵ���ʼλ��ʱ�˳�
    
    
    'strTmp = Replace(strSource, str_Count, "")
    If Len(strSource) > 0 Then str_Count = str_Count & strSource
    
    If InStr(strSource, Chr(&H5)) > 0 Or InStr(strSource, Chr(&H4)) > 0 Then
        strCmd = Chr(&H6)
    End If
    If InStr(strSource, Chr(&HD) & Chr(&HA)) > 0 Then
        strCmd = Chr(&H6)
    End If
        
    lngBegin = InStr(str_Count, Chr(&H5))
    If lngBegin > 0 Then
        lngEnd = InStr(lngBegin, str_Count, Chr(&H4))
    Else
        lngEnd = InStr(str_Count, Chr(&H4))
    End If
    
    Do While lngBegin > 0 And lngEnd > 0
    'ȡ��ǰһ�εĽ���������ķ���
        strLine = Replace(Mid(str_Count, lngBegin, lngEnd - lngBegin + 1), Chr(&HA), "")
        str_Count = Mid(str_Count, lngEnd + 1)
        
        '��ʼ����
        aData = Split(strLine, vbCr)
        
        '���ݳ��Ȳ���ʱ�˳�
        For i = LBound(aData) To UBound(aData)
            If aData(i) Like Chr(&H2) & "?P|?|*" Then
                aItem = Split(aData(2), "|")                '����
                strData = Mid(aItem(3), 1, 4) & "-" & Mid(aItem(3), 5, 2) & "-" & Mid(aItem(3), 7, 2) & " "                'ʱ��
                strData = strData & Mid(aItem(3), 9, 2) & ":" & Mid(aItem(3), 11, 2) & ":" & Mid(aItem(3), 13, 2)
                strData = strData & "|" & Val(aItem(2)) & "| |��|0"                 '�걾��
            
            ElseIf aData(i) Like Chr(&H2) & "?R|?|*" Or aData(i) Like "R|*" Then
                '�������
                aItem = Split(aData(i), "|")
                If UBound(aItem) >= 5 Then
                    strTmp = Replace(aItem(5), vbLf, "")
                    If InStr(strTmp, "^") > 0 Then strTmp = Split(strTmp, "^")(0)
                    strData = strData & "|" & aItem(3) & "|" & strTmp
                End If
            End If
        Next
        '�滻Ӣ��Ϊ����
        strData = Replace(strData, "CLEAR", "͸��")
        strData = Replace(strData, "CLOUDY", "����")
        strData = Replace(strData, "NEGATIVE", "����")
        strData = Replace(strData, "POSITIVE", "����")
        
        strData = Replace(strData, "TRACE-LYSED", "΢��")
        strData = Replace(strData, "TRACE-INTACT", "΢��")
        strData = Replace(strData, "TRACE", "������")
        
        strData = Replace(strData, "LT. YELLOW", "����ɫ")
        strData = Replace(strData, "DK. YELLOW", "���ɫ")
        strData = Replace(strData, "YELLOW", "��ɫ")
        
        strData = Replace(strData, "LT. RED", "ǳ��ɫ")
        strData = Replace(strData, "DK. RED", "���ɫ")
        strData = Replace(strData, "RED", "��ɫ")
        
        
        strData = Replace(strData, "DK. ORANGE", "���ɫ")
        strData = Replace(strData, "LT. ORANGE", "ǳ��ɫ")
        strData = Replace(strData, "ORANGE", "��ɫ")
        
        strData = Replace(strData, "DK. BROWN", "���ɫ")
        strData = Replace(strData, "BROWN", "��ɫ")
        
        strData = Replace(strData, "LT. GREEN", "ǳ��ɫ")
        strData = Replace(strData, "GREEN", "��ɫ")
        If strData <> "" Then strResult = strResult & "||" & strData
        
        lngBegin = InStr(str_Count, Chr(&H5))
        If lngBegin > 0 Then
            lngEnd = InStr(lngBegin, str_Count, Chr(&H4))
        Else
            lngEnd = InStr(str_Count, Chr(&H4))
        End If
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    'strCmd = Chr(&H6)
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





