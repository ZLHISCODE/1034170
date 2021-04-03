VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_DSFI1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:��΢��  DS-FI1 NIKON-55i
'����:��(NIKON)
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
    
    Dim aResults() As String
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream, strLine As String

    Dim str���� As String, str������Ŀ As String, str�걾�� As String
    Dim strFilePath As String, strFilename As String
     
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    strFilename = Dir(strFilePath & "\*.jpg")
    ReDim aResults(0)
    
    Do While strFilename <> ""
        If strFilename Like "########_*_*_*" Then
            str���� = Mid(strFilename, 1, 8)
            str���� = Mid(str����, 1, 4) & "-" & Mid(str����, 5, 2) & "-" & Mid(str����, 7, 2)
            
            str�걾�� = Trim(Split(strFilename, "_")(1))
            str������Ŀ = Trim(Split(strFilename, "_")(2))
            If IsDate(str����) And Val(str�걾��) <> 0 And str������Ŀ <> "" Then
                
                If CDate(str����) >= dtStart And CDate(str����) <= dtEnd Then
                    If lngSampleNO <> -1 Then
                        If lngSampleNO = Val(str�걾��) Then
                            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                            aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |Ѫ��|0|" & str������Ŀ & "|��" & vbCrLf & "��;110;0;" & strFilePath & "\" & strFilename
                        End If
                    Else
                        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |Ѫ��|0|" & str������Ŀ & "|��" & vbCrLf & "��;110;0;" & strFilePath & "\" & strFilename
                    End If
                End If
            End If
        End If
        strFilename = Dir
    Loop
    ResultFromFile = aResults
End Function

