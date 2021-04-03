VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_File_Tecan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ø����
'����:�ۿ�
'����:�¶�
'�û�:��������ҽԺ

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

    Dim strDate As String, strItem As String, strNO As String, strValue As String, strOD As String, strCoutOff As String, strSCO As String
    Dim strFilePath As String, strFileName As String
    Dim arrLine As Variant
    Dim lngLoop As Long, intB As Integer, intK As Integer
    Dim lngNo As Long, intCount As Integer
    
    On Error GoTo hErr
    
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    lngLoop = DateDiff("d", dtStart, dtEnd)
    ReDim aResults(0)

    For intCount = 0 To lngLoop
        strDate = Format(DateAdd("d", intCount, dtStart), "yyyy-MM-dd")
        strFileName = Dir(strFilePath & "\*_" & strDate & "_*.asc")
        
        Do While strFileName <> ""
            
            If objFile.FileExists(strFilePath & "\" & strFileName) Then
                Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFileName)
                lngNo = 0
                Do While Not objStream.AtEndOfStream
                    strLine = UCase(objStream.ReadLine)
                    arrLine = Split(strLine, Chr(9))
                    lngNo = lngNo + 1
                    If UBound(arrLine) >= 5 Then
                        strItem = arrLine(0)
                        strNO = arrLine(1)
                        If strNO Like "??-??-???" Then
                            
                            intB = Val(Split(strNO, "-")(0))
                            intK = Val(Split(strNO, "-")(1))
                            
                            strNO = intK + (intB - 1) * 16
                            
                            strValue = Trim("" & arrLine(2))
                            If strValue = "" Then strValue = "-"
                            strCoutOff = arrLine(3)
                            strOD = arrLine(4)
                            
                            If Val(strCoutOff) <> 0 Then
                                strSCO = Format(Val(strOD) / Val(strCoutOff), "0.000")
                            Else
                                strSCO = 0
                            End If
                            
                            If strValue <> "" And strDate <> "" And strNO <> "" Then
                                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                                If IsNumeric(strSCO) And IsNumeric(strCoutOff) Then
                                    aResults(UBound(aResults)) = strDate & "|" & strNO & "| |Ѫ��|0|" & strItem & "|" & strValue & "^" & strOD & "^" & strCoutOff & "^" & strSCO
                                Else
                                    aResults(UBound(aResults)) = strDate & "|" & strNO & "| |Ѫ��|0|" & strItem & "|" & strValue
                                End If
                            End If
                        End If
                    End If
                Loop
                
                objStream.Close
                Set objStream = Nothing
                
                '���˾Ͱ��ļ��ŵ�BAK��
                If Not objFile.FolderExists(strFilePath & "\Bak") Then Call objFile.CreateFolder(strFilePath & "\Bak")
                objFile.CopyFile strFilePath & "\" & strFileName, strFilePath & "\Bak\" & strFileName, True
                objFile.DeleteFile strFilePath & "\" & strFileName, True
                
            End If
            
            strFileName = Dir
        Loop
    Next
    ResultFromFile = aResults
    Exit Function
hErr:
    WriteLog "Tecan", Err.Number, Err.Description
End Function

Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    
    'iSendStep����ǰ���Ͳ����Ĳ��衣�����������ͽ������򷵻�0�������Զ���1���ء�
    'strResponse����ǰ��������Ӧ����
    'strSampleInfo�����͵ı걾��Ϣ��
    '   �걾֮����||�ָ�
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ���
    '   ��5�����أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��


    SendSample = ""
   
        
End Function

Private Function WritTxt(ByVal strFileName As String, strInfo As String)
    Dim objFileSystem As New FileSystemObject
    Dim objStream As TextStream
    
    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strInfo)
    objStream.Close
    Set objStream = Nothing
End Function

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

