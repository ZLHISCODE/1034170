VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ATB_1525"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ATB_1525΢���������
'����:
'����:������
'�û�:������ҽԺ

Public Function GetStartCmd() As String
    '��һ������ʱ����
End Function

Public Function GetEndCmd() As String
    '�ӿڹرյ�ʱ�����
End Function

Public Function GetAnswerCmd() As String
    '��ʱ���͵�ָ��
End Function

Public Function ResultFromFile(ByVal strfile As String, ByVal lngSampleNO As Long, _
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsAnti As ADODB.Recordset
    Dim strBh As String, strAnti As String
    Dim lngResultCount As Long
    strFilePath = objFile.GetParentFolderName(strfile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    On Error GoTo hErr
    
    With cnAccess
        .ConnectionString = "DBQ=" & strfile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};pwd=scsl"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "Select a.reqNo as ����id,a.samNo as �걾��,b.ExeDate as ����ʱ��,b.GermId as ϸ�� From ResultBill a,ResultExe b Where Format(b.ExeDate,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "'"
        strSQL = strSQL & "And a.reqNo = b.reqNo" & " And Format(b.ExeDate,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
        IIf(lngSampleNO = -1, "", " And val(a.samNo)=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        
        'rsTmp.MoveFirst
        Do Until rsTmp.EOF
            If strBh <> rsTmp("����id") Then '�±걾
                
                If Nvl(rsTmp("ϸ��")) = "" Then GoTo NextDo:
                
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(rsTmp("����ʱ��"), "yyyy-mm-dd hh:mm:ss") '����ʱ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val(rsTmp("�걾��")) '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "| "
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    IIf(IsNull(rsTmp("ϸ��")), " ", rsTmp("ϸ��")) '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|0"
                    
                strBh = rsTmp("����id")
                strAnti = ""
                
                strSQL = "Select antiId,micExplain as ҩ�����,iif(testNo='K-B','3',iif(testNo='DISK','2',iif(testNo='MIC','1'))) as ����,resShow From ResultMic Where reqNo=" & strBh
                Set rsAnti = cnAccess.Execute(strSQL)
                Do Until rsAnti.EOF
                    strAnti = strAnti & "|" & rsAnti!antiId & "KSS|" & rsAnti!ҩ����� & "^" & rsAnti!���� & "^" & rsAnti!resShow
                    rsAnti.MoveNext
                Loop
                aResults(lngResultCount) = aResults(lngResultCount) & strAnti
            End If
NextDo:
            rsTmp.MoveNext
        Loop
    End With
    '������ɣ��ر��ļ�
    If rsTmp.EOF Then
        cnAccess.Close
    End If
    ResultFromFile = aResults
    Exit Function
hErr:
    WriteLog "clsLisDev_ATB_1525", Err.Description, ""
End Function

