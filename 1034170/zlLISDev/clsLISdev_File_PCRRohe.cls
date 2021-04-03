VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_File_PCRRohe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:���� PCR
'����:����
'����:�¶�
'�û�:����úҵ������ҽԺ
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
    Dim str�걾�� As String, str������ As String, blnReturn As Boolean, str��Ŀ As String, str��� As String, strTmp As String
    Dim lngRow As Long, varLine As Variant
    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    
    If InStr(strFile, "-") <= 0 Then
        WriteLog "File_PCRRohe", "������ļ���,�ļ�����ʽ�� ��Ŀ-����", ""
        Exit Function
    Else
        str��Ŀ = Replace(Trim(Split(strFile, "-")(0)), strFilePath & "\", "")
        If str��Ŀ = "" Then
            WriteLog "File_PCRRohe", "������ļ���,�ļ�����ʽ�� ��Ŀ-����", ""
            Exit Function
        End If
    End If
    
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        varLine = Split(strLine, Chr(&H9))
        If lngRow = 0 Then
            str���� = Format(Now, "yyyy-MM-dd")
        Else
            If UBound(varLine) = 6 Then
                str�걾�� = Trim(Replace(varLine(3), "Sample", ""))
            
                If Val(str�걾��) > 0 Then
                    '"Ǧ����� ǦŨ�� ϡ�ͱ��� ѪǦ����
                    str��� = Replace(Trim(varLine(5)), "]", "")
                    str��� = Replace(str���, "[", "")
                    If InStr(str���, "E") > 0 Then
                        strTmp = Mid(str���, InStr(str���, "E") + 1)
                        If Len(strTmp) = 1 Then strTmp = "0" & strTmp
                        str��� = Mid(str���, 1, InStr(str���, "E")) & "+" & strTmp
                    End If
                    If str��� <> "" Then
                        str������ = "|" & str��Ŀ & "|" & str���
                    Else
                        str������ = "|" & str��Ŀ & "|<1.00E+03"
                    End If
                    If str������ <> "" Then
                        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| | |0" & str������
                    End If
                End If
            End If
        End If
        lngRow = lngRow + 1
        DoEvents
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_PCRRhoe.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function



