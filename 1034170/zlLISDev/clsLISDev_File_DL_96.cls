VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_DL_96"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:DL_96
'����:�϶�
'����:����
'�û�:XXXXXXXXXXX
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
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim str�걾�� As String, str��Ŀ As String
    Dim intLoop As Integer
    Dim intItem As Integer
    Dim aLine() As String
    Dim strItem As String
    
    
    On Error GoTo errH
    
    Dim rsErs As ADODB.Recordset
    Dim str���� As String

    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        .CursorLocation = adUseClient
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        
        If .State = adStateClosed Then Exit Function
        
        strSQL = "select * " & _
                 " from ϸ������ "

        Set rsTmp = cnAccess.Execute(strSQL)
        
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        ReDim aResults(0)
        
        Do While Not rsTmp.EOF
                
            str���� = Format("" & Now, "yyyy-MM-dd HH:mm:ss")
            str�걾�� = Val("" & rsTmp("�걾���"))
            
            For intLoop = 1 To 5
                If rsTmp("ϸ��" & intLoop) <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        
                    aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |" & _
                        Mid(rsTmp("ϸ��" & intLoop), 1, InStr(rsTmp("ϸ��" & intLoop), "(") - 1) & "|0"
                    If rsTmp("ҩ��" & intLoop) <> "" Then
                        aLine = Split(rsTmp("ҩ��" & intLoop), ";")
                        strItem = ""
                        For intItem = 0 To UBound(aLine) Step 6
                            If intItem + 6 <= UBound(aLine) Then
                                strItem = strItem & "|" & aLine(intItem + 1) & "|"
                                Select Case aLine(intItem + 4)
                                    Case "��ҩ"
                                        strItem = strItem & "R"
                                    Case "�н�"
                                        strItem = strItem & "I"
                                    Case "����"
                                        strItem = strItem & "S"
                                End Select
                                strItem = strItem & "^1^" & aLine(intItem + 3)
                            End If
                        Next
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & strItem
                    Else
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|1^R^"
                    End If
                End If
            Next
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    For intLoop = 0 To UBound(aResults)
        Debug.Print aResults(intLoop)
    Next
    
    Exit Function
errH:
    
    WriteLog "File_DL_96", Err.Description, ""
End Function



'0769-26208501
