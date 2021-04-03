VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LBY"
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
        If dtStart < date Then 'Ҫ��ѯ����
            strSQL = "SELECT a.Bh, Format(A.Date,'yyyy-mm-dd') As Rq, a.Bg, 'ѪҺ' As �걾, 0 As �ʿ�,b.Xh,b.Csz,zb" & _
                " FROM Brzl a, Brzl01 b WHERE a.Bh=b.Bh" & _
                IIf(lngSampleNO = -1, " And Format(A.Date,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                    " And Format(A.Date,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'", _
                " And a.Bh='" & Format(dtStart, "yymmdd") & "-" & Format(lngSampleNO, "000") & "'")
        End If
        If dtEnd >= date Then 'Ҫ�鵱��
            If Len(strSQL) > 0 Then blnUnion = True 'ҪUnion All
            
            strSQL = IIf(Len(strSQL) = 0, "", strSQL & " Union All ") & _
                "SELECT a.Bh, Format(A.Date,'yyyy-mm-dd') As Rq, a.Bg, 'ѪҺ' As �걾, 0 As �ʿ�,b.Xh,b.Csz,zb" & _
                " FROM TodayBrzl a, TodayBrzl01 b WHERE a.Bh=b.Bh" & _
                IIf(lngSampleNO = -1, " And Format(A.Date,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                    " And Format(A.Date,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'", _
                " And a.Bh='" & Format(dtStart, "yymmdd") & "-" & Format(lngSampleNO, "000") & "'")
        End If
        strSQL = strSQL & IIf(blnUnion, " Order By Rq Desc ,Bh,Xh", " Order By Format(A.Date,'yyyy-mm-dd') Desc ,a.Bh,b.Xh")

        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        Do While Not rsTmp.EOF
            If strBh <> rsTmp("Bh") Then '�±걾
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(rsTmp("Rq"), "yyyy-mm-dd hh:mm:ss") '����ʱ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val(Mid(rsTmp("Bh"), 8)) '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    IIf(IsNull(rsTmp("Bg")), " ", rsTmp("Bg")) '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    IIf(IsNull(rsTmp("�걾")), " ", rsTmp("�걾")) '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val(rsTmp("�ʿ�")) '�ʿ�
                    
                strBh = rsTmp("Bh")
            End If
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                IIf(IsNull(rsTmp("Xh")), "0", rsTmp("Xh")) & "|" & Val(IIf(IsNull(rsTmp("Csz")), "0", rsTmp("Csz"))) '���
            
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function