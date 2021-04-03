VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AYW9003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ȫ�Զ���Ѫ������ AYW9003
'����:�Ͼ�����Ƽ��������޹�˾
'����:������
'�û�:�˳����ۿ�ҽԺ
'����:2012-12-18

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
    Dim strBH As String, strAnti As String, str��� As String
    Dim lngResultCount As Long, str_��Ŀ As String

    On Error GoTo errH
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};pwd=508"
    cnAccess.Open
    If cnAccess.State = adStateClosed Then Exit Function

    strSQL = "Select A.biaobenhao as �걾��,A.jianyanriqi as ����,A.jianyanshijian as ʱ��, A.pt as PT, A.ptr AS ptr ,A.inr as inr,A.aptt as aptt,A.tt as TT,A.pt_fib AS FIB "
    strSQL = strSQL & " From bingrenziliao  A "

    strSQL = strSQL & " Where  A.jianyanriqi >=Cdate('" & Format(dtStart, "yyyy-MM-dd") & "')"
    strSQL = strSQL & " And  A.jianyanriqi < Cdate('" & Format(dtEnd + 1, "yyyy-MM-dd") & "') " & _
             IIf(lngSampleNo = -1, "", " And val(A.biaobenhao)=" & lngSampleNo) & " order by 1"
    WriteLog "AYW9003", strSQL, ""

    strBH = "***"
    ReDim aResults(0)

    Set rsTmp = cnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        cnAccess.Close
        WriteLog "AYW9003", "�޼�¼��", ""
        Exit Function
    End If
    Do Until rsTmp.EOF
        
        If strBH <> rsTmp("�걾��") Then  '�±걾
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)

            aResults(UBound(aResults)) = Format(Trim("" & rsTmp("����")) & " " & Trim("" & rsTmp("ʱ��")), "yyyy-mm-dd hh:mm:ss") '����ʱ��
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                                         Val(rsTmp("�걾��"))    '�������
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                                         "ѪҺ"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
            strBH = rsTmp("�걾��")
        End If

        If strBH = rsTmp("�걾��") Then
            If Trim("" & rsTmp("pt")) <> "" Then
                If str_��Ŀ = "" Then
                    str_��Ŀ = "PT|" & Replace(Replace(Trim("" & rsTmp("pt")), "��", ""), "��", "")
                Else
                    str_��Ŀ = str_��Ŀ & "|PT|" & Replace(Replace(Trim("" & rsTmp("pt")), "��", ""), "��", "")
                End If
            End If
            If Trim("" & rsTmp("ptr")) <> "" Then
                If str_��Ŀ = "" Then
                    str_��Ŀ = "PTR|" & Replace(Replace(Trim("" & rsTmp("ptr")), "��", ""), "��", "")
                Else
                    str_��Ŀ = str_��Ŀ & "|PTR|" & Replace(Replace(Trim("" & rsTmp("ptr")), "��", ""), "��", "")
                End If
            End If
            If Trim("" & rsTmp("inr")) <> "" Then
                If str_��Ŀ = "" Then
                    str_��Ŀ = "INR|" & Replace(Replace(Trim("" & rsTmp("inr")), "��", ""), "��", "")
                Else
                    str_��Ŀ = str_��Ŀ & "|INR|" & Replace(Replace(Trim("" & rsTmp("inr")), "��", ""), "��", "")
                End If
            End If
            If Trim("" & rsTmp("aptt")) <> "" Then
                If str_��Ŀ = "" Then
                    str_��Ŀ = "APTT|" & Replace(Replace(Trim("" & rsTmp("aptt")), "��", ""), "��", "")
                Else
                    str_��Ŀ = str_��Ŀ & "|APTT|" & Replace(Replace(Trim("" & rsTmp("aptt")), "��", ""), "��", "")
                End If
            End If
            If Trim("" & rsTmp("tt")) <> "" Then
                If str_��Ŀ = "" Then
                    str_��Ŀ = "TT|" & Replace(Replace(Trim("" & rsTmp("tt")), "��", ""), "��", "")
                Else
                    str_��Ŀ = str_��Ŀ & "|TT|" & Replace(Replace(Trim("" & rsTmp("tt")), "��", ""), "��", "")
                End If
            End If
            If Trim("" & rsTmp("FIB")) <> "" Then
                If str_��Ŀ = "" Then
                    str_��Ŀ = "FIB|" & Replace(Replace(Trim("" & rsTmp("fib")), "��", ""), "��", "")
                Else
                    str_��Ŀ = str_��Ŀ & "|FIB|" & Replace(Replace(Trim("" & rsTmp("fib")), "��", ""), "��", "")
                End If
            End If
        End If
        If str_��Ŀ <> "" Then
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & str_��Ŀ
            str_��Ŀ = ""
        End If

        WriteLog "AYW9003", aResults(UBound(aResults)), ""
        rsTmp.MoveNext
    Loop

    ResultFromFile = aResults
    Exit Function
errH:
    WriteLog "AYW9003", CStr(Erl()) & ":" & Err.Description, ""
End Function



