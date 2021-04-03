VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_EH_2060B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:EH_2060B �����
'����:����
'����:�¶�
'�û�:�����еڶ�����ҽԺ
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
    
    Dim aResults() As String
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream, strLine As String

    Dim str���� As String, str������Ŀ As String, str�걾�� As String, strImg As String
    Dim strFilePath As String, strFilename As String
    Dim lngS As Long, lngE As Long
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    strFilename = UCase(Dir(strFilePath & "\*.txt"))
    ReDim aResults(0)
    
    Do While strFilename <> ""
        str���� = "": str�걾�� = "": str������Ŀ = "": strImg = ""
        If UCase(strFilename) Like "###########.TXT" Then
            If objFile.FileExists(strFilePath & "\" & strFilename) Then
                Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFilename)
                Do While Not objStream.AtEndOfStream
                    strLine = objStream.ReadLine
                    
                    If strLine Like "[[]P07[]]*:*" Then
                        str�걾�� = Split(strLine, ":")(1)
                    ElseIf strLine Like "[[]P08[]]*:*" Then
                        str���� = Split(strLine, ":")(1)
                    ElseIf strLine Like "[[]P[1-4]#[]]*:*" Then
                        If Split(strLine, ":")(1) <> "" Then
                            lngS = InStr(strLine, "[") + 1
                            lngE = InStr(strLine, "]")
                            
                            str������Ŀ = str������Ŀ & "|" & Mid(strLine, lngS, lngE - lngS) & "|" & Split(strLine, ":")(1)
                        End If
                    ElseIf strLine Like "[[]P61[]]*:*" And UBound(Split(strLine, ":")) > 1 Then
                        strImg = strImg & "^ͼ1;110;0;" & Split(strLine, ":")(1) & ":" & Split(strLine, ":")(2)
                    ElseIf strLine Like "[[]P62[]]*:*" And UBound(Split(strLine, ":")) > 1 Then
                        strImg = strImg & "^ͼ2;110;0;" & Split(strLine, ":")(1) & ":" & Split(strLine, ":")(2)
                    End If
                Loop
                
                If Val(str�걾��) > 0 And IsDate(str����) And str������Ŀ <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |Ѫ��|0" & str������Ŀ
                    If strImg <> "" Then aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & Mid(strImg, 2)
                End If
                '���˾Ͱ��ļ��ŵ�BAK��
                objStream.Close
                Set objStream = Nothing
                If Not objFile.FolderExists(strFilePath & "\Bak") Then Call objFile.CreateFolder(strFilePath & "\Bak")
                objFile.CopyFile strFilePath & "\" & strFilename, strFilePath & "\Bak\" & strFilename, True
                objFile.DeleteFile strFilePath & "\" & strFilename, True
            End If
        End If
        strFilename = Dir
    Loop
    ResultFromFile = aResults
End Function



