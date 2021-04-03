VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LightCycleII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'����:ӫ�ⶨ��PCR������ LightCycleII
'����:���� Roche
'����:�¶�
'�û�:����ҽԺ
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
    Dim strChenel As String, str����Ա As String, str���� As String
    Dim strFilePath As String
    
    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str�걾�� As String, str������ As String, blnReturn As Boolean, i As Integer
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
        If strLine Like "*Run By:*" Then
            '����Ա
            str����Ա = Trim(Mid(strLine, InStr(strLine, "Run By:") + 7))
            
        ElseIf strLine Like "Run Date:*,*" Then
            'ʱ��
            str���� = Replace(Split(strLine, vbTab)(0), "Run Date: ", "")
            str���� = Replace(str����, ", ", " ")
            varTmp = Split(str����, " ")
            str���� = varTmp(2) & "-" & TrunMonth(varTmp(0)) & "-" & varTmp(1)
            If IsDate(str����) = False Then
                str���� = Format(Now, "yyyy-MM-dd")
            End If
        ElseIf InStr(strLine, vbTab) > 0 Then
            varItem = Split(strLine, vbTab)
            If UBound(varItem) = 5 Then
                str�걾�� = varItem(1)
                str������ = ""
                If Not (InStr(str�걾��, "-") > 0 Or _
                   InStr(str�걾��, "+") > 0 Or _
                   InStr(str�걾��, "/") > 0 Or _
                   InStr(str�걾��, "ST") > 0 Or _
                   InStr(str�걾��, "CONTROL") > 0) Then
                   
                    If str�걾�� Like "HB*" And Trim(varItem(3)) <> "" Then
                        str�걾�� = Mid(str�걾��, 3)
                        str������ = "HB|" & Trim(varItem(3))
                        
                        If Len(str�걾��) >= 10 Then
                            str���� = Format(CDate(Mid(str�걾��, 2, 2) & "-" & Mid(str�걾��, 4, 2) & "-" & Mid(str�걾��, 6, 2)), "yyyy-MM-dd")
                        End If
                        
                    ElseIf InStr(str�걾��, " ") > 0 And Trim(varItem(3)) <> "" Then
                        varTmp = Split(str�걾��, " ")
                        If IsNumeric(Mid(varTmp(1), 1, 1)) Then
                            str�걾�� = varTmp(1)
                        Else
                            str�걾�� = Mid(varTmp(1), 2)
                            
                            For i = 1 To Len(str�걾��)
                                If IsNumeric(Left(str�걾��, 1)) Then
                                    Exit For
                                Else
                                    str�걾�� = Mid(str�걾��, i + 1)
                                End If
                            Next
                        End If
                        str������ = varTmp(0) & "|" & Trim(varItem(3))
                        
                        If Len(str�걾��) >= 10 Then
                            str���� = Format(CDate(Mid(str�걾��, 2, 2) & "-" & Mid(str�걾��, 4, 2) & "-" & Mid(str�걾��, 6, 2)), "yyyy-MM-dd")
                        End If
                    End If
                    
                    If str������ <> "" Then
                        blnReturn = False
                        If lngSampleNO = -1 Then
                            'ȡָ�����ڼ�¼
                            If CDate(str����) >= dtStart And CDate(str����) <= dtEnd Then
                                blnReturn = True
                            End If
                        Else
                            If lngSampleNO = str�걾�� Then
                                'ָ���˱걾��
                                blnReturn = True
                            End If
                        End If
                        
                        If blnReturn Then
                            If str�걾�� <> "" Then
                                If IsDate(str����) = False Then str���� = Format(Now(), "yyyy-MM-dd")
                                If str����Ա = "" Then str����Ա = " "
                                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                                aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "|" & str����Ա & "|Ѫ��|0|" & str������
                            
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_LightCycleII.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function


Private Function TrunMonth(ByVal strMonth As String) As String
    'ת�����ĵ��·�Ϊ����
    Select Case UCase(strMonth)
        Case "һ��": TrunMonth = "01"
        Case "����": TrunMonth = "02"
        Case "����": TrunMonth = "03"
        Case "����": TrunMonth = "04"
        Case "����": TrunMonth = "05"
        Case "����": TrunMonth = "06"
        Case "����": TrunMonth = "07"
        Case "����": TrunMonth = "08"
        Case "����": TrunMonth = "09"
        Case "ʮ��": TrunMonth = "10"
        Case "ʮһ": TrunMonth = "11"
        Case "ʮ����": TrunMonth = "12"
        
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