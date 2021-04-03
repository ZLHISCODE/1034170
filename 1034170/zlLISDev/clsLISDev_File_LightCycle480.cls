VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LightCycle480"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ӫ�ⶨ��PCR������ LightCycle480
'����:���� Roche
'����:�µ���
'�û�:����������ҽԺ
'����:2012-09-12

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'��׼���������ļ���ȡ������
'   strFile������·�����ļ���
'   lngSampleNO���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
'   dtStart����ʼʱ�䡣���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
'   dtEnd������ʱ�䡣ֻ�е�ѡȡ����걾��lngSampleNO=-1��ʱ���ò�����Ч�������ָ���ò���ֵ����ʱ�䷶ΧΪ>=dtStart��
'   ���صļ���������(������������밴���±�׼��֯���)
'   Ԫ��֮����|�ָ�
'   ��0��Ԫ�أ�����ʱ��
'   ��1��Ԫ�أ��������
'   ��2��Ԫ�أ�������
'   ��3��Ԫ�أ��걾
'   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
'   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
'       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    
    Dim aResults() As String, objFile As New Scripting.FileSystemObject
    Dim strChenel As String, str���� As String, cTDH As String
    Dim strFilePath As String
    
    Dim objStream As TextStream
    Dim strLine As String, varLine As Variant
    Dim str�걾�� As String, str������ As String, lngRow As Integer
    
    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If (Not strLine Like "*Experiment:*") And (Not strLine Like "*Include*Color*") Then
          varLine = Split(strLine, vbTab)
          If UBound(varLine) >= 5 Then
             strChenel = Trim(varLine(3))      'HC1120706188
             strChenel = Trim(Mid(strChenel, 3))
             If IsNumeric(strChenel) Then
                str�걾�� = strChenel  '1120706188
                str���� = "20" & Mid(strChenel, 2, 2) & "-" & Mid(strChenel, 4, 2) & "-" & Mid(strChenel, 6, 2) & " " & Format(Now, "HH:MM:SS")
                If aResults(UBound(aResults)) <> "" Then
                    ReDim Preserve aResults(UBound(aResults) + 1)
                End If
                str������ = Trim(varLine(5))
                If IsDate(CDate(str����)) And str�걾�� <> "" And str������ <> "" Then
                    cTDH = "HBV"
                    If InStr(varLine(3), "HB") Then
                       cTDH = "HBV"
                    ElseIf InStr(varLine(3), "HC") Then
                       cTDH = "HCV"
                    End If
                    aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |Ѫ��|0|" & cTDH & "|" & str������
                End If
             End If
          End If
        End If
        DoEvents
    Loop
    
    For lngRow = 0 To UBound(aResults)
      Debug.Print aResults(lngRow)
    Next
    
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_LightCycle480.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function




