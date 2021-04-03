VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LG_PABER_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ѪС��ۼ���Ѫ���ӷ�����(LG-PABER-I)
'����:
'����:�¶�
'�û�:�Ϻ�һ��ҽԺ

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
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strBh As String
    
    On Error GoTo hErr
    
    ReDim aResults(0)
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    If Dir(strFile) = "" Then
        ResultFromFile = aResults
        Exit Function
    End If
    
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then
            ResultFromFile = aResults
            Exit Function
        End If
                
        strSQL = "select RecIndex as ����,TestName as ��Ŀ, unit as ��λ,SecResult as ���,MedicalRecordID as �걾��" & _
                " From Result " & _
                " Where unit<>'%' And SecResult<>0 And RecIndex>=" & Format(dtStart, "yyMMdd000") & _
                " and RecIndex<=" & Format(dtEnd, "yyMMdd999") & _
                IIf(lngSampleNO = -1, "", " And val(MedicalRecordID) = " & lngSampleNO) & _
                " order by Recindex ,MedicalRecordID"
        
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            ResultFromFile = aResults
            Exit Function
        End If
 
        strBh = "***"
        Do While Not rsTmp.EOF
            If strBh <> rsTmp("�걾��") And rsTmp("�걾��") <> 0 Then '�±걾
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                aResults(UBound(aResults)) = "20" & Mid(rsTmp("����"), 1, 2) & "-" & Mid(rsTmp("����"), 3, 2) & "-" & Mid(rsTmp("����"), 5, 2)     '����ʱ��
                
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val("" & rsTmp("�걾��"))   '�������
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "  '������
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|ѪҺ"  '�걾
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & IIf(Trim("" & rsTmp!��Ŀ) = "", Trim("" & rsTmp!��λ), Trim("" & rsTmp!��Ŀ)) & "|" & Format(rsTmp!���, "0.00")
                strBh = rsTmp("�걾��")
            ElseIf strBh = rsTmp("�걾��") And rsTmp("�걾��") <> 0 Then
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & IIf(Trim("" & rsTmp!��Ŀ) = "", Trim("" & rsTmp!��λ), Trim("" & rsTmp!��Ŀ)) & "|" & Format(rsTmp!���, "0.00")
            End If
            
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    Exit Function
hErr:
    WriteLog "File_Lg_Paber_I", CStr(Erl()) & "," & Err.Description, strSQL
End Function
