VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_MB5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:΢��Ԫ�ط����� MB5,MG2
'����:
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
    Dim strDate As String, strFilePath As String
    
    Dim objStream As TextStream, strLine As String, varItem As Variant
    Dim strNO As String, strItem As String, i As Integer
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
        varItem = Split(strLine, Chr(&H9))
        If UBound(varItem) >= 11 Then
            If IsNumeric(varItem(0)) Then
                    strDate = Format(Now, "yyyy-MM-dd HH:mm:ss")
                    strNO = varItem(0)
                    strItem = ""
                    For i = 3 To 11
                        If Trim$("" & varItem(i)) <> "" Then strItem = strItem & "|" & Format(i, "000") & "|" & Trim$("" & varItem(i))
                    Next
                    If strItem <> "" Then
                        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        aResults(UBound(aResults)) = strDate & "|" & strNO & "| |Ѫ��|0" & strItem
                    End If
            End If
        ElseIf UBound(varItem) >= 8 Then
            If IsNumeric(varItem(0)) Then
                    strDate = Format(Now, "yyyy-MM-dd HH:mm:ss")
                    strNO = varItem(0)
                    strItem = ""
                    For i = 4 To 8
                        If Trim$("" & varItem(i)) <> "" Then strItem = strItem & "|" & Format(i - 3, "000") & "|" & Trim$("" & varItem(i))
                    Next
                    If strItem <> "" Then
                        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        aResults(UBound(aResults)) = strDate & "|" & strNO & "| |Ѫ��|0" & strItem
                    End If
            End If
        ElseIf UBound(varItem) >= 5 Then
            If IsNumeric(varItem(0)) Then
                    strDate = Format(Now, "yyyy-MM-dd HH:mm:ss")
                    strNO = varItem(0)
                    strItem = ""
                    For i = 4 To 5
                        If Trim$("" & varItem(i)) <> "" Then strItem = strItem & "|" & Format(i + 2, "000") & "|" & Trim$("" & varItem(i))
                    Next
                    If strItem <> "" Then
                        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        aResults(UBound(aResults)) = strDate & "|" & strNO & "| |Ѫ��|0" & strItem
                    End If
            End If
        End If
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "MB5", strFile & " " & strLine, Err.Number & " " & Err.Description
End Function
