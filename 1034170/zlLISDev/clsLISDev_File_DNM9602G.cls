VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_DNM9602G"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:ø����DNM9602G
'ҽԺ:��������ҽԺ(¡��������ҽԺ)
'����:���
'ʱ��:2009/10/16
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsItem As ADODB.Recordset
    Dim lngResultCount As Long
    Dim blnUnion As Boolean
    Dim lngSampleID As Long
    Dim blnNew As Boolean           '�Ƿ��±걾��
    Dim lngSampleNumber As Long     '�걾ID
    Dim lngloop As Long
    Dim intLoop As Integer
    Dim strBatch As String          '��¼����
    Dim strResult As String         '���
    Dim strOD As String             'ԭʼ���(OD)
    Dim strCO As String             '�ж�ֵ(Cut Off)
    Dim strItem As String           '��Ŀ��ͨ����)
    Dim strBh As String             '��ǰ���
    Dim strCalc As String           '����ֵ
    Dim dateVerify As Date          '����ʱ��
    Dim StrTime As String
    Dim lng_count As Long
    Dim arrTmp(95) As Integer
    Dim i As Integer, j As Integer, k As Integer
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    lng_count = -1
    ResultFromFile = aResults
    
    '����ڷű걾
    For i = 0 To 7
        For j = 0 To 11
            arrTmp(i * 12 + j) = j * 8 + i
            'Debug.Print arrTmp(j)
        Next
    Next
    i = 0: j = 0
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        strSQL = "select * from �������˱� a , bzsz b where a.��� = b.bh  and " & _
                 " Format(b.rq,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "'" & _
                 " And Format(b.rq,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
                " order by ������,���,�׺�λ��"
        Set rsTmp = cnAccess.Execute(strSQL)

        Do Until rsTmp.EOF
            If lngSampleNumber <> rsTmp("������") Then
                lngSampleNumber = rsTmp("������")
                intLoop = 1
                
                If k > 0 And i < 95 And j < 95 And rsTmp!PLH = strBatch Then
                    j = j + 1
                    i = i + 1
                Else
                    i = 0
                    j = 0
                End If
                strBatch = rsTmp!PLH
                k = 1
            End If
            If strBh <> rsTmp("���") Then
                intLoop = 1
            End If


            '�õ���ǰOD
            strSQL = "select distinct * from bzsj where bh = '" & rsTmp("bh") & "'"

            Set rsItem = cnAccess.Execute(strSQL)

            If rsItem.EOF = False Then
                strOD = Nvl(rsItem("SJ" & arrTmp(i)))               '(Val(Replace(rsTmp("���"), "S", "")) - 1)
            End If

            '�õ���ǰCut Offֵ
            strItem = Nvl(rsTmp("xm" & intLoop & "c"))
            strSQL = "select * from xxsm where bh = '" & rsTmp("bh") & "' and xmbh = '" & strItem & "'"
            Set rsItem = cnAccess.Execute(strSQL)
            If rsItem.EOF = False Then
                strCO = Nvl(rsItem("co"))
            End If

            '����ֵ(��ʽ: OD/Cut Off
            If Val(strCO) <> 0 Then
                strCalc = Round(Val(strOD) / Val(strCO), 3)
            Else
                strCalc = 0
            End If
            '����ʱ��
            dateVerify = Nvl(rsTmp("rq")) & " " & Nvl(rsTmp("time"))

            '������
            strSQL = "select * from ������ where bh = '" & rsTmp("bh") & "'"
            Set rsItem = cnAccess.Execute(strSQL)
            If rsItem.EOF = False Then
                strResult = Nvl(rsItem("a" & arrTmp(j)))              '(Val(Replace(rsTmp("���"), "S", "")) - 1))
                
                If InStr(strResult, "-") > 0 Then
                    strResult = "����(-)"
                ElseIf InStr(strResult, "+") > 0 Then
                    strResult = "����(+)"
                Else
                    strResult = ""
                End If
            End If

            '�õ���ǰ����Ŀͨ����
            strItem = Nvl(rsTmp("xm" & intLoop & "c"))
            strSQL = "select * from xmcs where xmbh = '" & strItem & "'"
            Set rsItem = cnAccess.Execute(strSQL)
            If rsItem.EOF = False Then
                strItem = Nvl(rsItem("xmmc1"))
            End If

            If strItem <> "" Then
                ReDim Preserve aResults(lng_count + 1)
                aResults(UBound(aResults)) = dateVerify & "|" & rsTmp("������") & "| |Ѫ��|0|" & strItem & "|" & _
                                        strResult & "^" & strOD & "^" & strCO & "^" & strCalc
                WriteLog "���� IMS 972", aResults(UBound(aResults)), ""
            End If
            lng_count = lng_count + 1
            intLoop = intLoop + 1
            strBh = rsTmp("���")
            lngSampleNumber = rsTmp("������")
            If rsTmp("rq") <> StrTime Then
                i = 0
                j = 0
            End If
            StrTime = rsTmp("rq")
            rsTmp.MoveNext
        Loop
        ResultFromFile = aResults
    End With
End Function

