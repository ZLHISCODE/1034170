VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_EasyLyte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:EASYLYTE
'����:MEDICA
'����:�µ���
'�û�:�����е���ҽԺ
'����:2012-05-03
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

    Dim iStrPoint As Integer, intSpot As Integer
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim lngSector As Long, lngCup As Long, SampleID As String
    Dim iPos As Long, mstrCode As String, dblValue As Double, str���� As String
    Dim strLog As String, strTmp As String
    Dim strLine As String, strRq As String, sFlag As String
    Dim lng_S As Integer, lng_E As Integer
    Dim aData() As String, str_Code As String
    Dim intloop As Integer
    Dim strItem As String
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": strData = "": str���� = ""
    If InStr(strSource, ";") = 0 Then strReserved = strSource: Exit Sub
    
    mstrCode = strSource
    
    sFlag = Mid(mstrCode, InStr(mstrCode, ";"), 10)
    
    mstrCode = Replace(mstrCode, "PATIENT'S NAME", Chr(&HB))
    mstrCode = Replace(mstrCode, sFlag, sFlag & Chr(&H1C))
    
    lng_S = InStr(mstrCode, Chr(&HB))
    lng_E = InStr(mstrCode, Chr(&H1C))
    
    mstrCode = TrimBlank(mstrCode)
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
        aData = Split(str_Code, Chr(&HD) & Chr(&HA))
        For intloop = 0 To UBound(aData)
          strTmp = aData(intloop)
          If InStr(strTmp, "SAMPLE") > 0 Then
             SampleID = Mid(strTmp, InStr(strTmp, "SAMPLE") + 7)
             SampleID = Trim(SampleID)
          ElseIf InStr(strTmp, "ID#") > 0 Then
             SampleID = Mid(strTmp, InStr(strTmp, "ID#") + 3)
             SampleID = Trim(SampleID)
          ElseIf Mid(strTmp, 1, 2) = "Na" Then
             strData = strData & "|" & Trim(Mid(strTmp, 1, 2)) & "|" & Trim(Mid(strTmp, 3, 6))
             strData = strData & "|" & Trim(Mid(strTmp, 10, 2)) & "|" & Trim(Mid(strTmp, 12, 6))
             strData = strData & "|" & Trim(Mid(strTmp, 18, 2)) & "|" & Trim(Mid(strTmp, 20, 6))
          ElseIf InStr(strTmp, ";") > 0 Then
             strTmp = Replace(strTmp, Chr(&H1C), "")
             strRq = "20" & Mid(strTmp, 8, 2) & "-" & TrunMonth(Mid(strTmp, 1, 3)) & "-" & Trim(Mid(strTmp, 5, 2)) & " " & Trim(Mid(strTmp, 11)) & ":" & Format(Now, "ss")
             strLine = "||" & Format(strRq, "yyyy-MM-dd HH:mm:ss") & "|" & Val(SampleID) & "| |Ѫ��|0" & strData
          End If
        Next
        If Val(SampleID) > 0 And strLine <> "" Then
            strResult = strResult & strLine
        End If
        lng_S = InStr(mstrCode, Chr(&HB))
        lng_E = InStr(mstrCode, Chr(&H1C))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    

End Sub

Public Function TrimBlank(ByVal sTxt As String) As String
    Do While InStr(sTxt, vbCrLf & vbCrLf) <> 0
        sTxt = Replace(sTxt, vbCrLf & vbCrLf, vbCrLf)
    Loop
    TrimBlank = sTxt
End Function

Private Function TrunMonth(ByVal strMonth As String) As String
    'ת��Ӣ�ĵ��·�Ϊ����
    Select Case Mid(UCase(strMonth), 1, 3)
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select

End Function
