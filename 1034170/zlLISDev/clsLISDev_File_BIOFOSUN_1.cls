VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_BIOFOSUN_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'�û�����Դ������ҽԺ
'�����ͺţ�΢����BIOFOSUN-II
'���ߣ�����
'ʱ�䣺2014/4/18
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
        Dim objSystemFile As New FileSystemObject, cnAccess As New ADODB.Connection, rsTmp As New ADODB.Recordset
        Dim textStream As textStream
        Dim strResults() As String
        Dim strSQL As String, lngRecourd As Integer, i As Integer
        Dim strFilePath As String, filename As String, strBth As String
        
        On Error GoTo h
        strFilePath = objSystemFile.GetParentFolderName(strFile)
        If Len(strFilePath) = 0 Then strFilePath = App.Path
        With cnAccess
            .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
            .Open
        End With
        If cnAccess.State = adStateClosed Then Exit Function
        strSQL = "select b.TestRecordNum,b.plateID,b.GroupNum,b.Concentration as ���,b.SIR,b.ref_Value,b.AntibioticNum,a.enum,a.value as ����ҩ��, " & _
                        "c.testRecordNum , c.specimenNum, c.IDtype, c.DateTimeStamp " & _
                        "from (select * from tblEnumAntibiotic_copy where id>=109 and id<>161 and id <> 163) a,tblResultGroups_copy b, tblTestRecord c where " & _
                        "a.enum=b.AntibioticNum and c.testRecordNum=b.testRecordNum  " & vbNewLine
        strSQL = strSQL & " and format(c.DateTimeStamp,'yyyy-mm-dd') >= " & "'" & Format(dtStart, "yyyy-mm-dd") & "'" & vbNewLine
        strSQL = strSQL & " and format(c.DateTimeStamp,'yyyy-mm-dd') <= " & "'" & Format(dtEnd, "yyyy-mm-dd") & "'" & vbNewLine
        strSQL = strSQL & IIf(lngSampleNO = -1, "", " and c.specimenNum= " & lngSampleNO) & vbNewLine
        strSQL = strSQL & " order by b.TestRecordNum , b.plateID"
        Set rsTmp = cnAccess.Execute(strSQL)
        strBth = "***"
        lngRecourd = -1
        ReDim Preserve strResults(0)
        strResults(0) = ""
        Do Until rsTmp.EOF
            If strBth <> rsTmp("specimenNum").Value Then '�����µı걾��
                lngRecourd = lngRecourd + 1
                ReDim Preserve strResults(lngRecourd)
                strResults(lngRecourd) = Format(rsTmp("DateTimeStamp").Value, "yyyy-mm-dd hh:mm:ss") & "|" & rsTmp("specimenNum").Value & "| |" '�걾��
                strResults(lngRecourd) = strResults(lngRecourd) & rsTmp("IDtype").Value & "|0|" 'ϸ��
                strResults(lngRecourd) = strResults(lngRecourd) & rsTmp("����ҩ��").Value & "|" & rsTmp("sir").Value & "^1^" & rsTmp("���").Value
                strBth = rsTmp("specimenNum").Value
            Else
                strResults(lngRecourd) = strResults(lngRecourd) & "|" & rsTmp("����ҩ��").Value & "|" & rsTmp("sir").Value & "^1^" & rsTmp("���").Value
            End If
            rsTmp.MoveNext
        Loop
        For i = LBound(strResults) To UBound(strResults)
            Debug.Print strResults(i)
        Next
        rsTmp.Close
        cnAccess.Close
        ResultFromFile = strResults
        Exit Function
h:
    'MsgBox Err.Description, vbCritical, "Information"
    WriteLog "File_BIOFOSUN", CStr(Erl()) & ":" & Err.Description, ""
    Err.Clear
End Function

