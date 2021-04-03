VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_MDK3200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Ѫ����MDK3200
'ҽԺ:�����Ժ
'����:����
'ʱ��:2007/11/27
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
    ReDim aResults(0)
    aResults(0) = ""
    
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strsql = "": blnUnion = False
        
        If dtStart < date Then 'Ҫ��ѯ����
            strsql = "SELECT b.bh AS SampleID,b.date,a.xh ,a.csz as valuation " & _
                    "FROM BRZL  b, BRZL01 a  where b.bh = a.bh and " & _
                    "  Format(b.date,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                    " And Format(b.date,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'" & _
                    " order by date desc,b.bh"
        Else
            strsql = "SELECT b.bh AS SampleID,b.date,a.xh ,a.csz as valuation " & _
                    "FROM BRZL  b, BRZL01 a  where b.bh = a.bh and " & _
                    "  Format(b.date,'yyyy-mm-dd')='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                    " order by date desc,b.bh"
        End If

        Set rsTmp = cnAccess.Execute(strsql)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        Do While Not rsTmp.EOF
            If strBh <> rsTmp("SampleID") Then '�±걾
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                aResults(UBound(aResults)) = Format(rsTmp("date"), "yyyy-mm-dd hh:mm:ss") '����ʱ��
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                    Val(Mid(rsTmp("SampleID"), 9)) '�������
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                    " " '������
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & " |" & _
                    "ѪҺ" '�걾
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                    "0" '�ʿ�
            End If
            strBh = rsTmp("SampleID")
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                rsTmp("xh") & "|" & rsTmp("valuation")
                            
            rsTmp.MoveNext
        Loop
    End With
    
    ResultFromFile = aResults
End Function