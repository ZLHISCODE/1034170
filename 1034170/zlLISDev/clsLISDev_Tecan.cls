VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Tecan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Tecan Sunrise
'����:Tecan(��ʿ�ۿ�)
'����:�¶�
'�û�:��������ҽԺ
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
    
    Dim varLine As Variant
    'û����ô��ʱ�˳�
    If UBound(Split(strSource, vbCr)) < 9 Then
        strReserved = strSource
        Exit Sub
    End If
    
    '��ʼ����
    strResult = ""
    For intloop = 1 To 8
        strLine = ""
        varLine = Split(Trim(Split(strSource, vbCr)(intloop)), " ")
        For i = LBound(varLine) To UBound(varLine)
            strLine = strLine & ";" & Format(Val(varLine(i)) / 1000, "0.000")
        Next
        strResult = strResult & Mid(strLine, 2) & "|"
    Next
    
    
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

Public Function CmdAnalyse(strCmdType As String) As String
    '����:  �ֽ�����Ϊ����ִ�е�����
    
'    strControl = Me.cbo����.Text & ";" & Me.cbo�ο�����.Text & ";" & Me.cbo���Ƶ��.Text & ";" & Me.txt���ʱ�� & _
'                 ";" & Me.cbo���巽ʽ.Text & ":" & Me.cbo�հ���ʽ.Text
    
    Select Case strCmdType
        Case "Begin"                                                    '��ʼ ����
            CmdAnalyse = "" 'Chr(&H2) & "TR0" & Chr(&H3) & "7" & Chr(&HD) & "|+|-"
        Case "End"                                                      '���� ����
            CmdAnalyse = Chr(&H2) & "TR1" & Chr(&H3) & "6" & Chr(&HD) & "|+|-"
        Case "Revert"                                                       '�ظ�����
            CmdAnalyse = "+" 'Chr(&H2) & "+" & Chr(&H3) & "*" & Chr(&HD) '+*
        Case "Play"                     '�������
            CmdAnalyse = Chr(&H2) & "SM" & Chr(&H3) & Chr(&H1F) & Chr(&HD) & "|+|D" '��SM ����Dʱ������ȡ��
        Case "ReadData"    'ȡ��
            CmdAnalyse = Chr(&H2) & "DA" & Chr(&H3) & Chr(&H4) & Chr(&HD) 'SM
        Case "Out"
            CmdAnalyse = Chr(&H2) & "TR0" & Chr(&H3) & "7" & Chr(&HD) & "|+|-"
        Case "In"
            CmdAnalyse = Chr(&H2) & "TR1" & Chr(&H3) & "6" & Chr(&HD) & "|+|-"
        '----------------------------------------------------------------------------------
        Case "1(405)"                   '����һ
            CmdAnalyse = Chr(&H2) & "SW405" & Chr(&H3) & "4" & Chr(&HD) & "|+"
        Case "2(450)"                   '������
            CmdAnalyse = Chr(&H2) & "SW450" & Chr(&H3) & "4" & Chr(&HD) & "|+"
        Case "3(492)"                   '������
            CmdAnalyse = Chr(&H2) & "SW492" & Chr(&H3) & ":" & Chr(&HD) & "|+"
        Case "4(620)"                   '������
            CmdAnalyse = Chr(&H2) & "SW620" & Chr(&H3) & "1" & Chr(&HD) & "|+"
        '----------------------------------------------------------------------------------
        Case "����"                     '���巽ʽ����
            CmdAnalyse = ""
        Case "����"                     '���巽ʽ��
            CmdAnalyse = Chr(&H2) & "GD450 620" & Chr(&H3) & "'" & Chr(&HD) & "|" & Chr(&H2) & "3"
        '----------------------------------------------------------------------------------
        Case "1X"                       '���Ƶ�� 1X
            CmdAnalyse = ""
        Case "2X"
            CmdAnalyse = ""
        Case "3X"
            CmdAnalyse = ""
        '----------------------------------------------------------------------------------
        Case "����"
            CmdAnalyse = ""
        Case "�Լ�"
            CmdAnalyse = ""
        Case Else
'            If Val(strCmdType) > 0 Then 'ʱ��
'                CmdAnalyse = Chr(&H2) & "SH2 " & Val(strCmdType) * 10 & Chr(&H3) & Chr(&HD) & "|+"
'            End If
    End Select
    
End Function

