VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_XH6020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:XH-6020
'����:��Ӫ��������
'����:�¶�
'�û�:���칫˾
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
    Dim strChenel As String 'ͨ��
    Dim str����Ա As String, str���� As String
    Dim str�걾�� As String, str������ As String
    
    Dim strFilePath As String '�ļ�·��
    Dim blnReturn As Boolean '�Ƿ����
    Dim objStream As TextStream
    Dim strLine As String 'һ������
    
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
        strLine = DelSapce(strLine)
        If UBound(Split(strLine, " ")) = 7 Then
            str�걾�� = Split(strLine, " ")(3)
            If Split(strLine, " ")(5) Like "######" Then
                str���� = Format(CDate(Mid(Split(strLine, " ")(5), 1, 2) & "-" & Mid(Split(strLine, " ")(5), 3, 2) & "-" & Mid(Split(strLine, " ")(5), 5, 2)), "yyyy-MM-dd")
            Else
                str���� = Format(Now, "yyyy-MM-dd")
            End If
            
            blnReturn = False
            If lngSampleNO = -1 Then
                If CDate(str����) >= dtStart And CDate(str����) <= dtEnd Then
                    blnReturn = True
                End If
            Else
                If lngSampleNO = str�걾�� Then
                    'ָ���걾�ţ����
                    blnReturn = True
                End If
            End If
            
            If blnReturn Then
                strChenel = Trim(Split(strLine, " ")(2))
                str������ = Trim(Split(strLine, " ")(4))
                If Val(str�걾��) > 0 And strChenel <> "" And str������ <> "" Then
                    If aResults(UBound(aResults)) <> "" Then
                        ReDim Preserve aResults(UBound(aResults) + 1)
                    End If
                    aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |Ѫ��|0" & "|" & strChenel & "|" & str������
                End If
            End If
        End If
        DoEvents
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_XH6020.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function

