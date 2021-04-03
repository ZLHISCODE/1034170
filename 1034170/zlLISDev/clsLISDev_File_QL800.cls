VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_QL800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:QL800΢��Ԫ�ط�����
'����:��������ҽ����е���޹�˾
'����:�µ���
'�û�:���ݸ��ױ���Ժ

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
    Dim strBh As String, str��� As String
    Dim lngResultCount As Long
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "SELECT ���,��������,пŨ��,��Ũ��,��Ũ��,þŨ��,��Ũ��,��Ũ��,ͭŨ��,ǦŨ��,��Ũ��,��Ũ�� From ���鱨�� "
        strSQL = strSQL & " Where Format(��������,'yyyy-MM-dd 00:00:00')>='" & Format(dtStart, "yyyy-MM-dd 00:00:00") & "'"
        strSQL = strSQL & " And Format(��������,'yyyy-MM-dd 23:59:59')<='" & Format(dtEnd, "yyyy-MM-dd 23:59:59") & "'" & _
        IIf(lngSampleNO = -1, "", " And val(���)=" & lngSampleNO)
        strSQL = strSQL & vbNewLine & "order by ���"
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        
        'rsTmp.MoveFirst
        Do Until rsTmp.EOF
            If strBh <> rsTmp("���") Then '�±걾
                
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(rsTmp("��������"), "yyyy-mm-dd") & " " & Format(Now, "hh:mm:ss") '����ʱ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val(rsTmp("���")) '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "| "
                aResults(lngResultCount) = aResults(lngResultCount) & "| "
                aResults(lngResultCount) = aResults(lngResultCount) & "|0"
                aResults(lngResultCount) = aResults(lngResultCount) & "|Zn_" & "|" & Val(IIf(IsNull(rsTmp!пŨ��), " ", rsTmp!пŨ��))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Fe_" & "|" & Val(IIf(IsNull(rsTmp!��Ũ��), " ", rsTmp!��Ũ��))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Ca_" & "|" & Val(IIf(IsNull(rsTmp!��Ũ��), " ", rsTmp!��Ũ��))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Mg_" & "|" & Val(IIf(IsNull(rsTmp!þŨ��), " ", rsTmp!þŨ��))
                aResults(lngResultCount) = aResults(lngResultCount) & "|P_" & "|" & Val(IIf(IsNull(rsTmp!��Ũ��), " ", rsTmp!��Ũ��))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Se_" & "|" & Val(IIf(IsNull(rsTmp!��Ũ��), " ", rsTmp!��Ũ��))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Cu_" & "|" & Val(IIf(IsNull(rsTmp!ͭŨ��), " ", rsTmp!ͭŨ��))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Pb_" & "|" & Val(IIf(IsNull(rsTmp!ǦŨ��), " ", rsTmp!ǦŨ��))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Mn_" & "|" & Val(IIf(IsNull(rsTmp!��Ũ��), " ", rsTmp!��Ũ��))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Cd_" & "|" & Val(IIf(IsNull(rsTmp!��Ũ��), " ", rsTmp!��Ũ��))
                strBh = rsTmp("���")
            End If
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function



