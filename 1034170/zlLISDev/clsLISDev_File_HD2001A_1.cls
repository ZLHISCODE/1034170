VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_HD2001A_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:����оƬ�����
'�ͺ�:HD2001A
'����:�Ϻ���������Ƽ����޹�˾
'����:���
'�û�:
'����:2009-08-21

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

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
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
    Dim strSql As String, rsTemp As ADODB.Recordset
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
        strSql = "select testdate as ����ʱ��,ID as �������,testdoctor as ������,biaobenClass as �걾, " & _
                 "CA19_9,NSE,CEA,CA242,Ferritin,Beta_HCG,AFP,Free_PSA,PSA,CA125,HGH,CA15_3 from TestData " & _
                 "as a where format(cdate(testdate),'yyyy-mm-dd')>= '" & Format(dtStart, "yyyy-mm-dd") & "' " & _
                 "and format(cdate(testdate),'yyyy-mm-dd')<= '" & Format(dtEnd, "YYYY-MM-DD") & "'" & _
                 IIf(lngSampleNo = -1, "", " and VAL(ID)=" & lngSampleNo)
        Set rsErs = cnErs.Execute(strSql)
        
        While Not rsErs.EOF
            '�����ļ��ṹ��ȡ������Ϣ
            If Val("" & rsErs!�������) <> 0 Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                '�� 0 ��Ԫ�أ�����ʱ��
                '�� 1 ��Ԫ�أ��������
                '�� 2 ��Ԫ�أ�������
                '�� 3 ��Ԫ�أ��걾
                '�� 4 ��Ԫ�أ��Ƿ��ʿ�Ʒ
                '�� 5 ��Ԫ�ؿ�ʼΪ��������ÿ����Ԫ�ر�ʾһ��������Ŀ
                '�磺�� 5i ��Ԫ��Ϊ������Ŀ���� 5i+1 ��Ԫ�ر�ʾ������
                aResults(UBound(aResults)) = Format(CDate(rsErs!����ʱ��), "yyyy-mm-dd")                        '����ʱ��
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val("" & rsErs!�������)        '�������
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "                                  '������
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & rsErs!�걾                      '�걾
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                                  '�Ƿ��ʿ�Ʒ
                
                '������Ŀ
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|CA19_9|" & Trim(rsErs!CA19_9)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|NSE|" & Trim(rsErs!NSE)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|CEA|" & Trim(rsErs!CEA)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|CA242|" & Trim(rsErs!CA242)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|Ferritin|" & Trim(rsErs!Ferritin)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|Beta_HCG|" & Trim(rsErs!Beta_HCG)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|AFP|" & Trim(rsErs!AFP)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|Free_PSA|" & Trim(rsErs!Free_PSA)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|PSA|" & Trim(rsErs!PSA)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|CA125|" & Trim(rsErs!CA125)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|HGH|" & Trim(rsErs!HGH)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|CA15_3|" & Trim(rsErs!CA15_3)
            End If

            rsErs.MoveNext
        Wend
    End If
    ResultFromFile = aResults
End Function