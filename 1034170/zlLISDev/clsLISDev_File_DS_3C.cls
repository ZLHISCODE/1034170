VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_DS_3C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:΢�����ط����� DS-3C
'����:ɽ����ʢ�����������޹�˾
'����:�¶�
'�û�:��ú����ҽԺ
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
    Dim blnUnion As Boolean, strImg As String, iField As Integer, i As Integer
    
    Dim cnErs As New ADODB.Connection
    Dim strFileErs As String, rsErs As ADODB.Recordset
    Dim strImgErs As String, str���� As String
    Dim strChar As String
    Dim intMaxValue As Integer
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    strFileErs = strFile
    
    If Trim(Dir(strFileErs)) <> "" Then
        With cnErs
        .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then strFileErs = ""
        End With
    Else
        strFileErs = ""
    End If
        
    lngResultCount = -1
    
    If strFileErs <> "" Then
        strsql = "SELECT val(Mid(A.PATNUM,9)) As ˳���, A.PATJYSJ as ����, " & _
                 "B.NAME as ��Ŀ, B.CONTENT as ��� " & _
                 " FROM PAT_INFO AS A,PAT_CONT As B Where A.PATNUM=B.PATNUM And " & _
                 " A.PATJYSJ>='" & Format(dtStart, "yyyyMMdd") & _
                 "' And A.PATJYSJ<='" & Format(dtEnd, "yyyyMMdd") & "'" & _
                IIf(lngSampleNO = -1, "", " And val(Mid(PATNUM,9))=" & lngSampleNO)
        Set rsErs = cnErs.Execute(strsql)
        Do Until rsErs.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            
            str���� = "" & rsErs("����")
            str���� = Mid(str����, 1, 4) & "-" & Mid(str����, 5, 2) & "-" & Mid(str����, 7, 2)
            aResults(lngResultCount) = Format(str����, "yyyy-MM-dd")  '����ʱ��
            
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                Val("" & rsErs("˳���"))   '�������
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                  " "  '������
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                  "ѪҺ"  '�걾
            aResults(lngResultCount) = aResults(lngResultCount) & "|0"
            
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & Trim("" & rsErs!��Ŀ) & "|" & Val("" & rsErs!���)

            rsErs.MoveNext
        Loop
    End If
        

    ResultFromFile = aResults
End Function

