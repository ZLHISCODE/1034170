VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_XD511"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:XD511ȫ�Զ������������
'�ͺ�:XD511
'����:�Ϻ�����ҽ�Ƽ������޹�˾
'����:���
'�û�:�Ϻ�����ҽԺ�޺�ΰ
'����:2010-05-12
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
    Dim cnSQL2000 As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsAnti As ADODB.Recordset
    Dim str���� As String, arr����() As String
    Dim str������ As String, str�걾�� As String
    Dim str�걾���� As String, str�������� As String
    Dim lngResultCount As Long
    Dim iField As Integer
    Dim aPic() As String
    Dim strPic As String
    Dim strPicPath As String               'ͼ�����·��
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    '�����ļ��õ��������̵������ļ����ڼ�ʦվ�Ĳ���ѡ����ָ���ļ�ʱ��ָ���������ļ��������ļ�������Ҫ�õ����������£�
    '[BDE_MSQL_XLB]
    'Database Name = th_xlb2035
    'SERVER Name = PC081128
    'USER Name = sa
    'PASSWORD=
    '--------------------------------------------------------------------------------------------------------------------

    '�������ļ�
    Call ReadConf(strFile, "SERVER NAME,DATABASE NAME,USER NAME,PASSWORD", str����)
    arr���� = Split(str����, ",")
    If UBound(arr����) <> 3 Then Exit Function
    
    With cnSQL2000
        .ConnectionString = _
        "Provider=SQLOLEDB.1;Initial Catalog=" & arr����(1) & ";Password=" & arr����(3) & _
        ";Persist Security Info=True;User ID=" & arr����(2) & ";Data Source=" & arr����(0)
        .Open
     
        If .State = adStateClosed Then Exit Function
'        strSQL = "Select smpno,smptype,daterpt,testvalues_text,c.img,c.prev,c.notes From urinerpts a, micslices b,micimages c where a.id=b.id and b.sid=c.sid" & _
'                " And daterpt>='" & Format(dtStart, "yyyy-MM-dd") & "'" & _
'                " And daterpt<='" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
'                IIf(lngSampleNO = -1, "", " And smpno=" & lngSampleNO)
           
        '�Է����ô�����ͼ�ο��п��ޣ�����Ŀǰ������ͼ��
        strSQL = "Select smpno,smptype,daterpt,testvalues_text From urinerpts " & _
        " Where daterpt>=Convert(DATETIME,'" & Format(dtStart, "yyyy-MM-dd") & " 00:00:00',21)" & _
        " And daterpt<=Convert(DATETIME,'" & Format(dtEnd, "yyyy-MM-dd") & " 23:59:59',21)" & _
        IIf(lngSampleNo = -1, "", " And smpno=" & lngSampleNo)
        
        
        
        Set rsTmp = cnSQL2000.Execute(strSQL)
'        If rsTmp.EOF Then
'            .Close
'            Exit Function
'        End If
'
        ReDim aResults(0)
        lngResultCount = 0

        Do Until rsTmp.EOF
            If Val("" & rsTmp("smpno")) > 0 Then '�±걾
                str�걾�� = Replace(Trim("" & rsTmp("smpno")), "-", "")
                str�걾���� = Trim("" & rsTmp("smptype"))
                If str�걾���� = "" Then str�걾���� = "��Һ"
                str�������� = Format(Trim(rsTmp("daterpt")), "yyyy-MM-dd hh:mm:ss") '����ʱ��
                str������ = Trim(ReadTestValue_Text("" & rsTmp("testvalues_text")))

                If Val(str�걾��) > 0 And str������ <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str�������� & "|" & str�걾�� & "| |" & str�걾���� & "|0" & str������
                End If

            End If
            rsTmp.MoveNext
        Loop
    End With
        
    ResultFromFile = aResults
End Function

Private Sub ReadConf(strIniFile As String, ByVal strItems As String, ByRef strReturn As String)
    '�������ļ�
    'strIniFile :ini�����ļ���
    'strItems   :Ҫ��ȡ���ֶΣ��磺DATABASE NAME=th_xlb�е� DATABASE NAME,����ֶ���,�ָ�
    'strReturn  :��ȡ�ķ���ֵ����:DATABASE NAME=th_xlb�е� th_xlb ���������ֵ��,�ָ�
    
    Dim arrLine
    Dim strLine As String

    Dim blnOpen As Boolean
    Dim objFilesys As New FileSystemObject
    Dim objStream As TextStream

    Dim intLoop As Long, intCount As Integer
    Dim arrItems As Variant
    
    On Error GoTo ErrHand
    
    '�������ļ��ж�ȡ�������ļ��е�����cotf
    intCount = UBound(Split(strItems, ","))
    ReDim arrItems(intCount, 2)
    For intLoop = 0 To intCount
        arrItems(intLoop, 0) = Split(strItems, ",")(intLoop)
    Next
    
    If objFilesys.FileExists(strIniFile) Then
        Set objStream = objFilesys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intLoop = 0 To intCount
                    If Trim(arrLine(0)) = Trim(arrItems(intLoop, 0)) Then
                        arrItems(intLoop, 1) = Trim(arrLine(1))
                        Exit For
                    End If
                Next
           End If
        Loop
        objStream.Close
        strReturn = ""
        For intLoop = 0 To intCount
            strReturn = strReturn & "," & arrItems(intLoop, 1)
        Next
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        blnOpen = False
    End If
    
    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub

Private Function ReadTestValue_Text(ByVal strTestvalues_text As String) As String
    Dim strPcode As String
    Dim strCode As String
    Dim strValue As String
    Dim strResult As String
    Dim var_line() As String
    Dim var_Item() As String
    Dim strLine As String
    Dim strItem As String
    Dim i As Integer
    Dim j As Integer
    
    var_line = Split(strTestvalues_text, ">" & Chr(&HA) & "<")
    
    For i = LBound(var_line) To UBound(var_line)
        strLine = var_line(i)
        If InStr(var_line(i), ";") Then
            var_Item = Split(var_line(i), ";")
        
            If UBound(var_Item) > 0 Then
                For j = LBound(var_Item) To UBound(var_Item)
                    strItem = Trim(var_Item(j))
                    If strItem Like "pcode*" Then strPcode = Trim(Mid(strItem, InStr(strItem, "=") + 1))
                    If strItem Like "code*" Then strCode = Trim(Mid(strItem, InStr(strItem, "=") + 1))
                    If strItem Like "value*" Then strValue = Trim(Mid(strItem, InStr(strItem, "=") + 1))
                Next
                
                If strPcode <> "root" And strValue <> "" Then
                        strResult = strResult & "|" & strCode & "|" & strValue
                End If
            End If
        End If
    Next
    
    ReadTestValue_Text = strResult
End Function

'��ȡͼ��
Public Function DrawImg(objField As Field, strFile As String) As Boolean
'    Open strToDiskFile For Binary As #iFree
'        If lLeftOver > 0 Then
'            byteData() = rsFile("Bin").GetChunk(lLeftOver)
'            Put #iFree, , byteData()
'            lblStep.Width = lblBack.Width * CLng(lLeftOver) / CLng(nColSize)
'            DoEvents
'        End If
'
'        For i = 1 To nBlocks
'            byteData() = rsFile("Bin").GetChunk(nBlockSize)
'            Put #iFree, , byteData()
'            DoEvents
'            lblStep.Width = lblBack.Width * CLng(lLeftOver + i * nBlockSize) / CLng(nColSize)
'            If Err.Number <> 0 Then GoTo EndHere
'        Next i
'    Close #iFree
    
    Dim lngFileSize As Long
    Dim arrBin() As Byte
    Dim intFile As Integer
    
    On Error GoTo errH
    
    intFile = FreeFile
    Open strFile For Binary As intFile
    
        lngFileSize = objField.ActualSize
        ReDim arrBin(lngFileSize - 1) As Byte
        arrBin() = objField.GetChunk(lngFileSize)
        Put intFile, , arrBin()
    Close intFile
    
    DrawImg = True
    Exit Function
errH:
    Close intFile
    Kill strFile

End Function

