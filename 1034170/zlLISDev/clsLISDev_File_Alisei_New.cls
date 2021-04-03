VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Alisei_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ø����
'����:alisei
'����:�¶�
'�û�:��̨ɽҽԺ,��ͷ����ҽԺ
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rs��Ŀ As ADODB.Recordset, rsʱ�� As ADODB.Recordset
    Dim lngResultCount As Long, strcutoff As String, str��Ŀ As String
    Dim str���� As String, str���Խ�� As String
    
    On Error GoTo errH
    strFilePath = objFile.GetParentFolderName(strFile)
    
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strSQL = "Select Format(A.Adate,'yyyy-MM-dd') As ����, ID As �걾��, B.Name As ��Ŀ, C.Posneg, C.Result, C.Od, B.Cutoff" & vbNewLine & _
                "From Wrkl A, Test B, Pozzetti C" & vbNewLine & _
                "Where A.Sessionid = C.Sessionid And A.Wrklindex = C.Wrklindex And C.Testindex = B.Testindex And" & vbNewLine & _
                "      C.Sessionid = B.Sessionid And Format(A.Adate, 'yyyy-MM-dd') >= '" & Format(dtStart, "yyyy-MM-dd") & _
                "' And Format(A.Adate, 'yyyy-MM-dd') <= '" & Format(dtEnd, "yyyy-MM-dd") & "'"
        strSQL = strSQL & IIf(lngSampleNO = -1, "", " And ID='" & lngSampleNO & "'")
        Set rsTmp = cnAccess.Execute(strSQL)

        
        ReDim aResults(0)
        Do Until rsTmp.EOF
            
            If Val("" & rsTmp!�걾��) > 0 Then
                'lngResultCount = lngResultCount + 1
'                ReDim Preserve aResults(lngResultCount)
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                lngResultCount = UBound(aResults)
                str���� = "" & rsTmp!����
                aResults(lngResultCount) = str����  '����ʱ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val("" & rsTmp!�걾��)   '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                   " "  '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                  "ѪҺ"  '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|0" '�ʿ�
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp!��Ŀ & "|" & rsTmp!Posneg '���
                aResults(lngResultCount) = aResults(lngResultCount) & "^" & Val("" & rsTmp!OD) & _
                                                                      "^" & Val("" & rsTmp!CutOff) & _
                                                                      "^" & Val("" & rsTmp!result)
            
            End If
            rsTmp.MoveNext
      Loop
    End With
    ResultFromFile = aResults
    Exit Function
errH:
    WriteLog "Alisei_new", Err.Number, Err.Description
End Function





