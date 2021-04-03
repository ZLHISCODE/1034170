VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_SM3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:SM-3�Զ���ø�������
'����:�����к�������ʯҽ����Ʒ������
'����:�¶�
'�û�:ɽ������ɽҽԺ

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
        Dim strBh As String
    
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
                
120         strSQL = "Select customerid as �걾��,Format(checkdate,'yyyy-MM-dd') as ����,kindsymbol as ��Ŀ,decision as ���" & _
                    " From platerecord " & _
                    " where checkdate>=#" & Format(dtStart, "yyyy-MM-dd") & "#" & _
                    " And checkdate<=#" & Format(dtEnd, "yyyy-MM-dd") & "#" & _
                    IIf(lngSampleNO = -1, "", " and id = " & lngSampleNO) & _
                    " Order By customerid,checktime "
                
122         Set rsTmp = cnAccess.Execute(strSQL)
124         If rsTmp.EOF Then
126             .Close
128             ResultFromFile = aResults
                Exit Function
            End If
 
130         strBh = "***"
132         Do While Not rsTmp.EOF
134             If strBh <> rsTmp("�걾��") And rsTmp("�걾��") <> 0 Then '�±걾
136                 If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
138                 aResults(UBound(aResults)) = Trim("" & rsTmp("����"))
                
140                 aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val("" & rsTmp("�걾��"))   '�������
142                 aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "  '������
144                 aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|ѪҺ"  '�걾
146                 aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
148                 aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Trim("" & rsTmp!��Ŀ) & "|" & IIf(Trim("" & rsTmp!���) = "", " ", Trim("" & rsTmp!���))
150                 strBh = rsTmp("�걾��")
152             ElseIf strBh = rsTmp("�걾��") And rsTmp("�걾��") <> 0 Then
154                 aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Trim("" & rsTmp!��Ŀ) & "|" & IIf(Trim("" & rsTmp!���) = "", " ", Trim("" & rsTmp!���))
                End If
            
156             rsTmp.MoveNext
            Loop
        End With
158     ResultFromFile = aResults
        Exit Function
hErr:
160     WriteLog "File_SM3", CStr(Erl()) & "," & Err.Description, strSQL
End Function




