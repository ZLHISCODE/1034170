VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_LBY_N6C"
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
    
    Dim iSampleNO As Integer, iPos As Long
    Dim aDatas() As String, i As Integer
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        If Not blnReadItem Then '����һ�����ݵĿ�ʼλ��
            iStrPoint = InStr(iStrPoint, strSource, "BB")
            '�ж�һ�����Ƿ�ʼ
            If iStrPoint > 0 Then
                '��¼һ���������
                iStrPoint = iStrPoint + 2
                iCurrGrpEntry = iStrPoint - 2
                '�����Ƿ�����
                iPos = InStr(iStrPoint, strSource, "B28")
                If iPos > 0 Then
                    '������ţ�����Ϊ����
                    iSampleNO = Val(Mid(strSource, iStrPoint + 8, 4))
                    '����ʱ��
                    'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                    iPos = InStr(iStrPoint + 26, strSource, "B")
                    iStrPoint = iPos - 8
                    strData = Format(ReadSubString(strSource, iStrPoint, 8), "yyyy-mm-dd hh:mm:ss")
                            
                    '������ţ�����Ϊ����
                    strData = strData & "|" & iSampleNO
                            
                    '������
                    strData = strData & "| " '��������֧��
                    '�걾
                    strData = strData & "|ѪҺ" '��������֧��
                    '�ʿ�Ʒ
                    strData = strData & "|0"
                    
                    '׼������
                    blnReadItem = True
                End If
            Else
                Exit Do
            End If
        Else
            '������һ���������
            iPos = InStr(iStrPoint, strSource, "BB")
            If iPos > 0 Then
                strBuffer = Mid(strSource, iStrPoint, iPos - iStrPoint)
                iStrPoint = iPos
            Else
                strBuffer = Mid(strSource, iStrPoint, Len(strSource) - iStrPoint + 1)
                iStrPoint = Len(strSource) + 1
            End If
            aDatas = Split(strBuffer, "B")
            For i = 1 To UBound(aDatas)
                strData = strData & "|" & Mid(aDatas(i), 1, 2)
                strData = strData & "|" & Val(Mid(aDatas(i), 3))
            Next
                
            strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                
            blnReadItem = False
            '�������ѽ���
            iCurrGrpEntry = 0
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