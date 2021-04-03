VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_QQ96"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:QQ-96��ѧ�������߷�����
'�ͺ�:QQ-96
'����:ʯ��ׯ�������Ƽ����޹�˾
'����:���
'�û�:����ͭ�ʵ���ҽԺ
'����:2010-01-19
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
    'strFile������·�����ļ���
    'lngSampleNo���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
    'dtStart����ʼʱ�䣬���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
    'dtEnd������ʱ�䡣ֻ��ѡȡ����걾(lngSample=-1)ʱ�ò�����Ч�������ָ���ò�������ʱ�䷶ΧΪ>=dtStart
    '���صļ��������飨������������밴�����±�׼��֯�����
    'Ԫ��֮���� | �ָ�
    '�� 0 ��Ԫ�أ�����ʱ��
    '�� 1 ��Ԫ�أ��������
    '�� 2 ��Ԫ�أ�������
    '�� 3 ��Ԫ�أ��걾
    '�� 4 ��Ԫ�أ��Ƿ��ʿ�Ʒ
    '�ӵ� 5 ��Ԫ�ؿ�ʼΪ��������ÿ����Ԫ�ر�ʾһ��������Ŀ
    '�磺�� 5i ��Ԫ��Ϊ������Ŀ���� 5i+1 ��Ԫ�ر�ʾ������
    
    Dim aResults() As String
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim cnErs As New ADODB.Connection, strFileErs As String, rsErs As ADODB.Recordset, strImgErs As String
    Dim str_������Ŀ As String
        
    strFilePath = objFile.GetParentFolderName(strFile)  'ȡ·��
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strFileErs = strFile
    
    If Trim(Dir(strFileErs)) <> "" Then
        With cnErs
            .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};"
            .Open
            If .State = adStateClosed Then strFileErs = ""
        End With
    Else
        strFileErs = ""
    End If
    
    ReDim aResults(0)
    strImgErs = ""
    
    If strFileErs <> "" Then
        strSQL = "select A.RQ As ����ʱ�� ,B.YBH As �������,'ѪҺ' as �걾, C.ENG AS ������Ŀ, A.userCon as ������ " & _
                 "From BL as A, PATIENT as B ,REAG as C " & _
                 "Where A.SampId = B.ID And C.ID = A.ReagId and Format(Cdate(A.RQ),'yyyy-mm-dd')>= '" & Format(dtStart, "yyyy-MM-dd") & "' " & _
                 "And Format(Cdate(A.RQ),'yyyy-mm-dd')<= '" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
                 IIf(lngSampleNO = -1, "", " and VAL(B.YBH)=" & lngSampleNO)
                 
        Set rsErs = cnErs.Execute(strSQL)
        
        While Not rsErs.EOF
            '�����ļ��ṹ��ȡ������Ϣ
            If Val("" & rsErs!�������) <> 0 Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                aResults(UBound(aResults)) = Format(CDate(rsErs!����ʱ��), "yyyy-mm-dd") & "|" & Val("" & rsErs!�������) & "| |" & _
                                        rsErs!�걾 & "|0|" & rsErs!������Ŀ & "|" & rsErs!������
                'WriteLog "���� KPS-I", aResults(UBound(aResults)), ""
            End If

            rsErs.MoveNext
        Wend
    End If
    ResultFromFile = aResults
End Function

