VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_BIOFOSUN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ϸ��������ҩ�������� BIOFOSUN
'����:
'����:�¶�
'�û�:������ҽԺ

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

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsSlav As ADODB.Recordset
    Dim i As Integer, strDate As String
    On Error GoTo errH
    
         
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd="
    cnAccess.Open
    If cnAccess.State = adStateClosed Then Exit Function
    strSQL = "SELECT A.specimenNum,A.isolatenum,A.testDate,A.LIS_ID_ID" & _
             " FROM LIS_MAST as A Where A.testDate >=#" & Format(dtStart, "yyyy-MM-dd 00:00:00") & "#" & _
             " And A.TestDate<=#" & Format(dtEnd, "yyyy-MM-dd 23:59:59") & "#"
    
    If lngSampleNo <> -1 Then
        strSQL = strSQL & " And  a.sepecimenNum = " & lngSampleNo
    End If
    WriteLog "File_BIOFOSUN", strSQL, ""
    
    ReDim aResults(0)
    
    Set rsTmp = cnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        cnAccess.Close
        WriteLog "File_BIOFOSUN", "�޼�¼��", ""
        Exit Function
    End If
        
    Do Until rsTmp.EOF
         
        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            
        aResults(UBound(aResults)) = Format(rsTmp!testDate, "yyyy-MM-dd")  '����ʱ��
        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(rsTmp("specimenNum")) '�������
        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| |" & rsTmp!LIS_ID_ID & "|0"
        strSQL = "SELECT A.Med_CODE,a.SIR,a.Concentration " & _
                 " FROM LIS_SLAVE as A where a.specimenNum = '" & rsTmp!specimenNum & "' and A.isolateNum='" & rsTmp!isolatenum & "'"
        strSQL = strSQL & " Order by A.med_group "
        
        Set rsSlav = cnAccess.Execute(strSQL)
        Do Until rsSlav.EOF
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & rsSlav!med_code & "|" & rsSlav!SIR & "^1^" & rsSlav!Concentration
            rsSlav.MoveNext
        Loop
        rsTmp.MoveNext
    Loop
        
    ResultFromFile = aResults
    Exit Function
errH:
    WriteLog "File_BIOFOSUN", CStr(Erl()) & ":" & Err.Description, ""
End Function





