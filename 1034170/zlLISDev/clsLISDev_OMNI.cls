VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_OMNI"
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
    
    Dim iStrPoint As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim iPos As Integer
    Dim strDate As String, lngSampleNO As Long, strItemName As String, dblItemValue As Double
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    iStrPoint = 1: blnReadItem = False
    strDate = "": lngSampleNO = -1
    Do While iStrPoint <= Len(strSource)
        If Not blnReadItem Then '���ұ걾������Ϣ
            strBuffer = ReadLine(strSource, iStrPoint)
            Do While Len(strBuffer) > 0
                Select Case True
                    Case InStr(UCase(strBuffer), UCase("Meas Date/Time")) > 0 'ʱ��
                        iPos = InStr(UCase(strBuffer), UCase("Meas Date/Time"))
                        strDate = Trim(Mid(strBuffer, iPos + 15, 21))
                        '���ڸ�ʽΪ��dd.mm.yyyy hh:mm
                        strDate = Format(CDate(Replace(strDate, ".", "/")), "yyyy-mm-dd hh:mm:ss")
                        
                        strBuffer = Mid(strBuffer, iPos + 36)
                    Case InStr(UCase(strBuffer), UCase("Patient ID")) > 0 '�걾��
                        iPos = InStr(UCase(strBuffer), UCase("Patient ID"))
                        lngSampleNO = Val(Mid(strBuffer, iPos + 13, 10))
                        
                        strBuffer = Mid(strBuffer, iPos + 23)
                    Case Else
                        strBuffer = ""
                End Select
            Loop
            If strDate <> "" And lngSampleNO > 0 Then '������Ϣ������ɣ���λ��������Ŀ��ʼλ��
                'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                strData = strDate
                '������ţ�����Ϊ����
                strData = strData & "|" & lngSampleNO
                '������
                strData = strData & "| " '��������֧��
                '�걾
                strData = strData & "|ѪҺ" '��������֧��
                '�ʿ�Ʒ
                strData = strData & "|0" '��������֧�֣���Ϊ���ʿ�Ʒ
                
                Do While iStrPoint <= Len(strSource)
                    strBuffer = ReadLine(strSource, iStrPoint)
                    If InStr(UCase(strBuffer), UCase("Meas. val. at")) > 0 Then Exit Do
                Loop
                blnReadItem = True
            End If
        Else '������Ŀ
            strBuffer = ReadLine(strSource, iStrPoint)
            Do While Len(strBuffer) > 0
                iPos = 1
                strItemName = ReadItemName(strBuffer, iPos)
                strBuffer = Replace(LTrim(Mid(strBuffer, iPos)), "%", "")
                
                If InStr(UCase(LABITEMS), "," & UCase(strItemName) & ",") > 0 Then '�Ϸ��ļ�����Ŀ
                    strData = strData & "|" & strItemName
                    dblItemValue = Val(strBuffer)
                    strData = strData & "|" & dblItemValue
                    
                    strBuffer = Replace(strBuffer, CStr(dblItemValue), "", 1, 1)
                End If
            Loop
        End If
    Loop
            
    strResult = strData
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

Private Function ReadLine(ByVal strSource As String, iPoint As Integer) As String
    '��ȡһ���ַ��������ַ�ָ���Ƶ���һ�п�ʼλ��
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, Chr(13) + Chr(10))
    If iPos = 0 Then iPos = Len(strSource) + 1
    
    ReadLine = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 2
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '��ȡһ��������Ŀ���ƣ������ַ�ָ���Ƶ���Ŀ����֮��
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function
