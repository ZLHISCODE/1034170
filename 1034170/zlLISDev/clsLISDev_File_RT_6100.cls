VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_RT_6100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'����: RT-6100
'����:�����׶�Raytoø�⹤��վRT-6100
'����:����
'�û�:�ɶ���һ����ҽԺ
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
    Dim strsql As String, rsTmp As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean
    Dim intLoop As Integer
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strsql = "": blnUnion = False
        strsql = "SELECT PlateID,SampleID,ItemName,ResultQualityName,ResultQuantity,resultabs,TestDate " & _
                "FROM Table_Plate_Result  a  where  " & _
                 IIf(lngSampleNO = -1, "  Format(TestDate,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                " And Format(TestDate,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'", _
                "  a.TestDate ='" & Format(dtStart, "yyyy-mm-dd") & "' and  SampleID='" & Format(lngSampleNO, "00#") & "'") & _
                " order by TestDate,SampleID "

        Set rsTmp = cnAccess.Execute(strsql)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        Do Until rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & Replace(rsTmp("testDate"), "/", "-")
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("SampleID") & "| | |0"
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("ItemName") & "|" & rsTmp("ResultQualityName") & "^" & rsTmp("ResultQuantity") & "^" & rsTmp("resultabs") & "^0"
            aResults(lngResultCount) = Mid(aResults(lngResultCount), 2)
            rsTmp.MoveNext
        Loop
    End With
    
    For intLoop = 0 To UBound(aResults)
        Debug.Print aResults(intLoop)
    Next
    ResultFromFile = aResults
End Function

