VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_C2000A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Ѫ���� C2000-A
'����:�������������޹�˾
'����:�¶�
'�û�:������ҽԺ��̫ԭ��

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
        Dim strSQL As String, rsTmp As ADODB.Recordset
        Dim strBh As String, strItem As String
        On Error GoTo hErr
    
100     ReDim aResults(0)
102     strFilePath = objFile.GetParentFolderName(strFile)
104     If Len(strFilePath) = 0 Then strFilePath = App.Path
    
106     If Dir(strFile) = "" Then
108         ResultFromFile = aResults
            Exit Function
        End If
110     With cnAccess
112         .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
114         .Open
116         If .State = adStateClosed Then
118             ResultFromFile = aResults
                Exit Function
            End If
120         strSQL = "  " & _
                    " Select Val(a.sample_id) as �걾��,a.sample_date as ����,b.assay_id as ͨ����,b.assay_result_1 as ���" & _
                    " from sample_info a,assay_info b" & vbNewLine & _
                    " Where a.sample_Key=B.sample_key And a.sample_id<>0 And a.sample_date>=#" & Format(dtStart, "yyyy-MM-dd") & "# And a.sample_date<=#" & Format(dtEnd, "yyyy-MM-dd") & "# " & _
                    IIf(lngSampleNo = -1, "", " And val(a.sample_id)=" & lngSampleNo) & " Order by a.sample_date,a.sample_id"
                
122         Set rsTmp = cnAccess.Execute(strSQL)
124         If rsTmp.EOF Then
126             .Close
128             ResultFromFile = aResults
                Exit Function
            End If
 
130         strBh = "***"
132         strItem = ""
134         Do While Not rsTmp.EOF
136             If strBh <> rsTmp("�걾��") & "_" & Format(rsTmp("����"), "yyyy-mm-dd") Then '�±걾
138                 If strItem <> "" Then
140                     If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
142                     aResults(UBound(aResults)) = strItem
                    End If
                
144                 strItem = Format(rsTmp("����"), "yyyy-mm-dd")  '����ʱ��
146                 strItem = strItem & "|" & Val("" & rsTmp("�걾��"))   '�������
148                 strItem = strItem & "| "  '������
150                 strItem = strItem & "|ѪҺ"  '�걾
152                 strItem = strItem & "|0"
154                 strItem = strItem & "|1" & rsTmp!ͨ���� & "|" & rsTmp!���
156                 strBh = rsTmp("�걾��") & "_" & Format(rsTmp("����"), "yyyy-mm-dd")
                Else
158                 strItem = strItem & "|1" & rsTmp!ͨ���� & "|" & rsTmp!���
                End If
160             rsTmp.MoveNext
            Loop
                
162         If strItem <> "" Then
164             If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
166             aResults(UBound(aResults)) = strItem
            End If
        End With
168     ResultFromFile = aResults
        Exit Function
hErr:
170     WriteLog "C2000a", CStr(Erl()) & ":" & Err.Description, ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    
End Sub



