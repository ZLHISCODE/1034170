VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_lumino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  Lumino
'ҽԺ:  ����
'����:  ����
'ʱ��:  2008/1/16
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
    Dim adata() As String, iDataPoint As Integer, aitem() As String
    
    Dim strChenal As String, dblValue As Double, iPos As Long, iStartPoint As Integer
    Dim lngloop As Long, intLoop As Integer
    Dim strDate As String
    Dim i As Integer
    Dim strTmp As String
    Dim strSampleNO As String           '�걾��
    Dim strBarcode As String            '�����
    Dim strImageData As String          'ͼ����Ϣ
    Dim strLine As String               '��ǰҪ�������ִ�
    Dim lngBegin As Long                '��ʼλ��
    Dim lngEnd As String                '����λ��
    Dim blnRun As Boolean               '�Ƿ�ִ��
    Dim intRow As Integer               '��ǰ����
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    
    
    'û���ҵ���ʼλ��ʱ�˳�
    If InStr(strSource, "?") <= 0 Then
        strReserved = strSource
        strCmd = Chr(&H6)
        Exit Sub
    End If
    
    'û���ҵ�����λ��ʱ�˳�
    adata = Split(strSource, vbCrLf)
    For intLoop = 0 To UBound(adata)
        '����Ľ��������´��ٽ���
        If intRow > 0 Then
            strReserved = strReserved & adata(intLoop) & vbCrLf
        End If
        '���ҽ���
        If adata(intLoop) = "" Then
            intRow = intLoop
        End If
    Next
    
    'û���ҵ�����λ��ʱ�˳�
    If intRow = 0 Then
        strReserved = strSource
        strCmd = Chr(&H6)
        Exit Sub
    End If
    
    '��ʼ����
    For intLoop = 0 To intRow
        
        If Trim(adata(intLoop)) <> "" Then
            'ȥ������Ŀո�
            strLine = adata(intLoop)
            For i = 2 To 10
                strLine = Replace(strLine, Space(i), Space(1))
            Next
            strLine = Replace(strLine, Space(2), Space(1))
            aitem = Split(strLine, Space(1))
            
            If UBound(aitem) <= 2 Then
                strSampleNO = aitem(1)
            Else
                strData = strData & "||" & Now & "|" & strSampleNO & "| |Ѫ��|0|" & _
                            Trim(aitem(1)) & "|" & Val(Trim(aitem(4)))
            End If
        End If
    Next
    
    strResult = Mid(strData, 3)
    strCmd = Chr(&H6)
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











