VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_MB_80"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�û��������е�һ����ҽԺ
'������MB-80
'mdb��ȡ��ʽ
'���ߣ�

Option Explicit

Public Function GetStartcmd() As String
'��һ������ʱ����
End Function

Public Function GetEndCmd() As String
'�ӿڹرյ�ʱ�����
End Function

Public Function GetAnswerCmd() As String
'��ʱ���͵�ָ��
End Function

Public Sub Analyse(ByVal strSource As String, _
                   strResult As String, _
                   strReserved As String, _
                   strCmd As String)
'���뺯��
End Sub

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
    Dim ConnSQL As New ADODB.Connection
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim lngResultCount As Long, strBh As String
    Dim str���� As String
    Dim str_Tmp���� As String
    Dim str_Tmp����2 As String
    Dim objFile As New Scripting.FileSystemObject
    Dim strFilePath As String
    Dim strTDM As String
    
    On Error GoTo errHandle
    strFilePath = objFile.GetParentFolderName(strFile)

    If Len(strFilePath) = 0 Then strFilePath = App.Path

    With ConnSQL
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        strSQL = ""
        strSQL = "select ��Ʒ�� as �걾��, Mid(��������,1,6) as ����1, ������Ŀ as ͨ����, ������ as ��� from eds a  " & _
                 "where " & _
                 "Mid(��������,1,6) " & ">= " & Format(dtStart, "YYMMDD") & "" & _
               " and Mid(��������,1,6)" & "<= " & Format(dtEnd, "YYMMDD") & IIf(lngSampleNO = -1, " ", " and ��Ʒ�� = " & lngSampleNO) & " order by ��Ʒ��, ��������  "


        Set rsSQL = ConnSQL.Execute(strSQL)

        If rsSQL.EOF Then
            .Close
            ReDim Preserve aResults(0)
            aResults(0) = ""
            ResultFromFile = aResults
            Exit Function
        End If

        '                If rsTmp.EOF Then
        '            .Close
        '            Exit Function
        '        End If

        lngResultCount = -1
        strBh = "***"
        Do While Not rsSQL.EOF
            If strBh <> rsSQL("�걾��") Then    '�±걾
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)

                str���� = rsSQL("����1")
                str���� = Mid(str����, 1, 2) & "-" & Mid(str����, 3, 2) & "-" & Mid(str����, 5, 2)  '������ڳ�����yyyy-mm-dd��
                aResults(lngResultCount) = Format(str����, "yyyy-mm-dd")   '����ʱ��

                'str���� = Format(rsSQL("ResultTime"), "yyyy-mm-dd")

                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                           Val(rsSQL("�걾��"))  '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                         " "  '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                           "ѪҺ"  '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|0"    '& _
                                                                              ' Val(rsSQL("�ʿ�")) '�ʿ�
                strBh = rsSQL("�걾��")
            End If

            strTDM = rsSQL("ͨ����")
            strTDM = Replace(strTDM, "G-֬����", "G1")
            strTDM = Replace(strTDM, "1-3-��-D�Ͼ���", "13BD")

            If strTDM <> "" Or IsNull(strTDM) = False Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                           Trim(strTDM) & "|" & IIf(IsNull(rsSQL("���")), "*", rsSQL("���"))    '���
            End If

            rsSQL.MoveNext

        Loop

    End With

    If rsSQL.EOF Then
        ConnSQL.Close
    End If
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "MB_80", "MB_80", Err.Number & " " & Err.Description
End Function

Public Function SendSample(strSampleInfo As String, _
                           iSendStep As Integer, _
                           blnSuccess As Boolean, _
                           Optional ByVal strResponse As String = "", _
                           Optional ByVal blnUndo As Boolean = False, _
                           Optional ByVal iType As Integer = 0) As String
'����˫��ָ���
End Function







