VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_rayto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strSQL = "": blnUnion = False
        strSQL = "SELECT Sample_ID as sampleid,item_test_date as sjdate,item_ename as testnumber,item_result_det as valuation " & _
                "FROM Table_Pa_Item  a  where  " & _
                 IIf(lngSampleNO = -1, "  Format(item_test_date,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                " And Format(item_test_date,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'", _
                " And a.Sample_ID='" & Format(dtStart, "yyyymmdd") & Format(lngSampleNO, "00") & "'") & _
                " order by item_test_date desc,Sample_ID"

        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        Do While Not rsTmp.EOF
            If strBh <> rsTmp("SampleID") & Format(rsTmp("sjdate"), "yyyy-mm-dd") Then '�±걾
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(rsTmp("sjdate"), "yyyy-mm-dd hh:mm:ss") '����ʱ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val(rsTmp("SampleID")) '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    " " '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    "ѪҺ" '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    "0" '�ʿ�
                    
                strBh = rsTmp("SampleID") & Format(rsTmp("sjdate"), "yyyy-mm-dd")
            End If
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                UCase(rsTmp("testnumber")) & "|" & rsTmp("valuation")
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function