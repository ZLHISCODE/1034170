VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Mejer600_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'����:Mejer600
'ҽԺ:XXXXXXXX
'����:����
'���ʱ��:20090710

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
    Dim blnEndSample As Boolean '�걾����
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean
    Dim aDatas() As String, strItemResult As String, i As Integer
    Dim strNow As String        '��ǰ�ܽ������������ִ�
    Dim aLine() As String
    Dim intLoop As Integer
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    
    'û���ҵ���ʼ����λ��ʱ�˳�
    If InStr(strSource, "") = 0 Or InStr(strSource, "") = 0 Then
        strReserved = strSource
        strCmd = Chr(&H6)
        Exit Sub
    End If
    strReserved = Mid(strSource, InStr(strSource, "") + 1)
    strNow = Mid(strSource, InStr(strSource, "") + 1, InStr(strSource, "") - 1)
    
    aLine = Split(strNow, vbCrLf)
    
    'ʱ��
    strData = Format(Trim(Mid(aLine(1), 8)) & " " & Trim(aLine(2)), "yyyy-mm-dd hh:mm:ss") & _
                "|" & Val(Mid(aLine(1), 5, 3)) & "| |��Һ|0"
    
    
    For intLoop = 4 To UBound(aLine)
        If Trim(aLine(intLoop)) <> "" And Trim(aLine(intLoop)) <> "" Then
            strData = strData & "|" & Trim(Mid(aLine(intLoop), 1, 4)) & "|" & Val(Mid(aLine(intLoop), 9)) & _
                IIf(Trim(Mid(aLine(intLoop), 6, 3)) = "", "", "(" & Trim(Mid(aLine(intLoop), 6, 3)) & ")")
        End If
    Next
    
    strData = Replace(strData, "Brown Yellow", "�ػ�ɫ")
    strData = Replace(strData, "Thin Yellow", "����ɫ")
    strData = Replace(strData, "Yellow", "��ɫ")
    strData = Replace(strData, "Green", "��ɫ")
    strData = Replace(strData, "Primrose", "����ɫ")
    strData = Replace(strData, "Red", "��ɫ")
    strData = Replace(strData, "Orange", "��ɫ")
    strData = Replace(strData, "Blue", "��ɫ")
    strData = Replace(strData, "Brown", "��ɫ")
    strData = Replace(strData, "Unselected", "δѡ")
    strData = Replace(strData, "Other", "����")
    strData = Replace(strData, "Filemot", "�ƺ�ɫ")
    strData = Replace(strData, "Little Dark", "��ģ��")
    strData = Replace(strData, "Dark", "ģ��")
    strData = Replace(strData, "Clear", "�峺")
    strData = Replace(strData, "Feculent", "����")
    strData = Replace(strData, "Negative", "����")
    strData = Replace(strData, "Normal", "����")
    strData = Replace(strData, "Positive", "����")
    strData = Replace(strData, "Small", "1+")
    strData = Replace(strData, "Moderate", "2+")
    strData = Replace(strData, "Large", "3+")
    strData = Replace(strData, "Trace", "΢��")
    strData = Replace(strData, "Thin", "ǳ��ɫ")
    
    If Len(strData) > 0 Then
        strResult = strData
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
    
    iPos = InStr(iPoint, strSource, Chr(13) + Chr(10))
    If iPos > 0 Then
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 2
    Else
        iPos = InStr(iPoint, strSource, Chr(13))
        If iPos = 0 Then iPos = InStr(iPoint, strSource, Chr(10))
        If iPos = 0 Then iPos = Len(strSource) + 1
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 1
    End If
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '��ȡһ��������Ŀ���ƣ������ַ�ָ���Ƶ���Ŀ����֮��
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function


