VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Omlipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Password As String
'����:���ڹ����ض����׷�����
'�ͺţ�Omlipo
'����:
'�û�:����������ҽԺ
'ʹ�÷���:��ȡmdb���ݿ��ļ�
'**        ��


Public Function GetStartCmd() As String
    '��һ������ʱ����
End Function

Public Function GetEndCmd() As String
    '�ӿڹرյ�ʱ�����
End Function

Public Function GetAnswerCmd() As String
    '��ʱ���͵�ָ��
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
    Dim strSQL As String, strFilePath As String
    Dim objFile As New Scripting.FileSystemObject
    Dim errdb As New ADODB.Connection
    Dim cnAccess As New ADODB.Connection, aResults() As String
    Dim rsTmp As ADODB.Recordset
    Dim strSample As String, ResultCount As Long
    Dim str���� As String, Study_Time As String
    Dim Series_ItemName As String, Series_Value As String
    
    
    On Error GoTo errHand
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    If Dir(strFile) = "" Then Exit Function
    Password = "gsOmlipo"
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";pwd=" & Password & ";Driver={Microsoft Access Driver (*.mdb)}"  '�����ݿ�
        .Open
        If errdb.State = adStateClosed Then
           strSQL = "SELECT a.csampleNo,a.cdateDiagnose ,b.iprojectInfo ,b.cresult  " & vbNewLine & _
                    "from gs_diagnoseResult a ,gs_diagnoseDetail b " & vbNewLine & _
                    "Where a.idiagnoseId=b.idiagnoseId " & _
                    " and Format(A.cdateDiagnose,'yyyy-mm-dd') Between " & "'" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                    " and " & "'" & Format(dtEnd, "yyyy-mm-dd") & "'" & _
                    IIf(lngSampleNO = -1, "", " and  val(a.csampleNo) =" & lngSampleNO) & " order by A.csampleNo"
           Set rsTmp = cnAccess.Execute(strSQL)
            
            If rsTmp.EOF Then
                .Close
                ReDim Preserve aResults(0)
                aResults(0) = ""
                ResultFromFile = aResults
            Exit Function
                
            End If
            
            ResultCount = -1
            strSample = "***"
            Do Until rsTmp.EOF
                If strSample <> rsTmp("csampleNo") Then
                ResultCount = ResultCount + 1
                ReDim Preserve aResults(ResultCount)
                aResults(ResultCount) = Format(rsTmp("cdateDiagnose") & " " & Study_Time, "yyyy-mm-dd hh:mm:ss") '����ʱ��
                str���� = Format(rsTmp("cdateDiagnose"), "yyyy-mm-dd")
                aResults(ResultCount) = aResults(ResultCount) & "|" & Val(" " & rsTmp("csampleNo")) '�걾���
                aResults(ResultCount) = aResults(ResultCount) & "|" & " " '������
                aResults(ResultCount) = aResults(ResultCount) & "|" & "Ѫ��" '�걾
                aResults(ResultCount) = aResults(ResultCount) & "|0" '�Ƿ��ʿ�
                
                
                
                strSample = rsTmp("csampleNo")
                End If
                If rsTmp("cresult") <> "" Then
                    aResults(ResultCount) = aResults(ResultCount) & "|" & _
                    rsTmp("iprojectInfo") & "|" & Replace(rsTmp("cresult"), "��", "")
                    
                End If
            rsTmp.MoveNext
            Loop
        End If
    End With
    ResultFromFile = aResults
    
    Exit Function
    
errHand:
    WriteLog "omlipo", Err.Number & ":" & Err.Description, ""
End Function


