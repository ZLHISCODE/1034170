VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_SE9000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'��ҽ��Ժ
'2007-07-09 �¶� �޸�
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
    
    Dim iStrPoint As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim strSampleType As String '��������
    Dim iItemNO As Integer
    Dim sngItemValue As Single
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    If strSource Like "*D**" Then
        iStrPoint = 1: strBuffer = "": strData = "": blnReadItem = False: strSampleType = "": iItemNO = 0: sngItemValue = 0
        Do While iStrPoint <= Len(strSource)
            If Not blnReadItem Then '����һ�����ݵĿ�ʼλ��
                iStrPoint = InStr(iStrPoint, strSource, "D1")
                '�ж�һ�����Ƿ�ʼ
                If iStrPoint > 0 Then
                    '��¼һ���������
                    iStrPoint = iStrPoint + 1
                    iCurrGrpEntry = iStrPoint - 1
                    '�����Ƿ�����
                    If Len(strSource) - iStrPoint >= 61 Then
                        '����ʱ��
                        iStrPoint = iStrPoint + 1
                        strSampleType = ReadSubString(strSource, iStrPoint, 1)   '��������
                        
                        strData = ReadSubString(strSource, iStrPoint, 2)   '��
                        strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '��
                        strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '��
                        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                        strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                                
                        '������ţ�����Ϊ����
                        iStrPoint = iStrPoint + 12
                        strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 13))
                                
                        '������
                        strData = strData & "| " '��������֧��
                        '�걾
                        strData = strData & "|ѪҺ" '��������֧��
                        '�ʿ�Ʒ
                        strData = strData & "|" & IIf(strSampleType = "C", 1, 0) 'U:���桢E:STAT������S:��׼���ߡ�C:�ʿ�
                        
                        '׼������
                        iStrPoint = iStrPoint + 28
                        blnReadItem = True: iItemNO = 0
                    End If
                Else
                    Exit Do
                End If
            Else
                strBuffer = ReadSubString(strSource, iStrPoint, 1)
                If strBuffer = "" Then '�������
                    strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                    
                    '���豸���͡������յ�������
                    strCmd = strCmd & Chr(&H6&)
                    
                    blnReadItem = False
                    '�������ѽ���
                    iCurrGrpEntry = 0
                Else
                    '������Ŀ
                    iStrPoint = iStrPoint - 1
                    '���±�������ĿΪ5λ
                    strBuffer = ReadSubString(strSource, iStrPoint, IIf(InStr(",0,13,14,15,16,17,", "," & CStr(iItemNO) & ",") > 0, 5, 4))
                    '��Ŀͨ��ֱ�ӱ�עΪAxx������xxΪ��Ŀ���
                    strData = strData & "|A" & iItemNO
                    'ȷ��С��λ��(����Ϊ5�ı���С��2λ��ѪС�������δ����ϸ��Ϊ�������ౣ��1λ)
                    sngItemValue = Val(strBuffer) / IIf(Len(strBuffer) = 5, 100, IIf(iItemNO = 7 Or iItemNO = 22, 1, 10))
                    Select Case iItemNO
                        Case 1 '��ϸ������
                            sngItemValue = sngItemValue / 10
                        Case 2 'Ѫ�쵰��
                            sngItemValue = sngItemValue * 10
                        Case 6 'ƽ����ϸ��Ѫ�쵰��Ũ��
                            sngItemValue = sngItemValue * 10
                        Case 22 '��ѪС�����
                            sngItemValue = sngItemValue / 10
                    End Select
                    strData = strData & "|" & sngItemValue
                    
                    iItemNO = iItemNO + 1: iStrPoint = iStrPoint + 1 '���������
                End If
            End If
        Loop
                
        If blnReadItem Then
            strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
            '�������ѽ���
            iCurrGrpEntry = 0
        End If
        '����δ������ԭʼ��
        If iCurrGrpEntry > 0 Then
            strReserved = Mid(strSource, iCurrGrpEntry)
        Else
            '���豸���͡������յ�������
            strCmd = strCmd & Chr(&H6&)
        End If
    Else
        strReserved = strSource
        strCmd = strCmd & Chr(&H6&)
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