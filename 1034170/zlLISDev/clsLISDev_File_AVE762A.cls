VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AVE762A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:AVE762A �����
'����:��ɳ����
'����:�¶�
'�û�:����������

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
        Dim strSQL As String, rsTmp As ADODB.Recordset, rsAnti As ADODB.Recordset
        Dim strBh As String, str��� As String, str�걾�� As String
        Dim lngResultCount As Long
        Dim str���� As String, arr����() As String
        Dim str���� As String, iField As Integer
        Dim aPic() As String
        Dim strPic As String
        Dim strPicPath As String               'ͼ�����·��
        Dim intZIP As Integer
        Dim i As Integer, strImg As String, strImgTitle As String, iCount As Integer
        On Error GoTo errH
100     strFilePath = objFile.GetParentFolderName(strFile)
102     If Len(strFilePath) = 0 Then strFilePath = App.Path
    
        '�����ļ��õ��������̵������ļ����ڼ�ʦվ�Ĳ���ѡ����ָ���ļ�ʱ��ָ���������ļ��������ļ�������Ҫ�õ����������£�
        '[BDE_MSQL_XLB]
        'Database Name = th_xlb2035
        'SERVER Name = PC081128
        'USER Name = sa
        'PASSWORD=
        '--------------------------------------------------------------------------------------------------------------------
    
104     strPicPath = ReadIni("AVE762A", "ͼƬ·��", App.Path & "\AVE762A.ini")
106     If strPicPath = "" Then strPicPath = "E:": Call WriteIni("AVE762A", "ͼƬ·��", "E:", App.Path & "\AVE762A.ini")
        '�������ļ�
108     Call ReadConf(App.Path & "\AVE762A.ini", "SERVER NAME,DATABASE NAME,USER NAME,PASSWORD", str����)
110     arr���� = Split(str����, ",")
112     If UBound(arr����) <> 3 Then Exit Function
    
        '�Ƿ�ʹ��ZIP��ʽѹ���ļ�(10.24.120)�����ʹ��
114     intZIP = Val(ReadIni("AVE762A", "�Ƿ�ѹ��", App.Path & "\AVE762A.ini", "1"))
116     Call WriteIni("AVE762A", "�Ƿ�ѹ��", CStr(intZIP), App.Path & "\AVE762A.ini")
    
118     With cnAccess
120         .ConnectionString = _
            "Provider=SQLOLEDB.1;Initial Catalog=" & arr����(1) & ";Password=" & arr����(3) & _
            ";Persist Security Info=True;User ID=" & arr����(2) & ";Data Source=" & arr����(0)
122         .Open

124         If .State = adStateClosed Then
126             WriteLog "Ave762A", "��125��", "δ�����ӵ�SQLSERVER"
                Exit Function
            End If
    '128         strSQL = "select A.id,A.chkdate,b.item0 as M1,b.item1 as M2,b.item2 as M3,b.item3 as M4,b.item4 as M5, " & _
    '                        "cast((CASE WHEN cast(b.item0 as float)=0 THEN 0 ELSE 100-(CAST(b.Item6 AS float)/ CAST(b.Item0 AS float)) * 100 END) as decimal(5,1)) AS M6," & _
    '                        "cast((CASE WHEN cast(b.item0 as float)=0 THEN 0 ELSE (cast(b.item6 AS float)/ cast(b.item0 AS float)) * 100 END) as decimal(5,1)) AS M7, " & _
    '                        " b.item7 as M8,b.item8 as  M9,b.item9 as M10,B.item10 as M11,B.item11 as M12,b.item12 as M13,b.item13 as M14,b.item14 as M15,b.item15 as M16," & _
    '                        "b.item16 AS M17,b.item17 AS M18,b.item18 AS M19,b.item19 as M20,b.item20 AS M21,B.iteM21 as M22,B.ITEM22 AS M23,B.ITEM23 AS M24," & _
    '                        "b.item24 AS M25,b.item25 AS M26 ,b.item26 AS M27,b.item27 AS M28,b.item28 AS ITEM29,b.item29 as M30,b.item30 as M31,b.item31 AS M32," & _
    '                        "b.item32 AS M33, b.item33 AS M34 , b.item34 AS M35 " & _
    '                        ",C.item0 as  C1,c.item1 as  C2,c.item2 as  C3,c.item3 as  C4,c.item4 as  C5,c.Item5 as  C6,c.item6 as  C7,c.item7 as  C8" & _
    '                        ",c.item8 as  C9,c.item9 as  C10,c.item10 as  C11,d.PicPrt " & vbNewLine & _
    '                        " from sick A LEFT OUTER JOIN " & _
    '                        "      chemic C ON A.ID = C.ID LEFT OUTER JOIN " & _
    '                        "      micro B  ON A.ID = B.ID LEFT OUTER JOIN " & vbNewLine & _
    '                        "      micro_out D ON A.id = d.id and a.MICROID = substring(d.id,9,4)   " & vbNewLine & _
    '                        " where  a.chkdate>='" & Format(dtStart, "yyyy.MM.dd") & "' and a.chkdate<='" & Format(dtEnd, "yyyy.MM.dd") & "'" & _
    '                        " order by A.chkdate,A.id"
    '                        "      micro_out D ON A.id = substring(d.Id,1,8) and a.MICROID = substring(d.id,9,4)   " & vbNewLine &
128                strSQL = "select A.id,A.chkdate,b.item0 as M1,b.item1 as M2,b.item2 as M3,b.item3 as M4,b.item4 as M5, " & _
                        " cast((CASE WHEN cast(b.item0 as float)=0 THEN 0 ELSE 100-(CAST(b.Item6 AS float)/ CAST(b.Item0 AS float)) * 100 END) as decimal(5,1)) AS M6," & _
                        " cast((CASE WHEN cast(b.item0 as float)=0 THEN 0 ELSE (cast(b.item6 AS float)/ cast(b.item0 AS float)) * 100 END) as decimal(5,1)) AS M7," & _
                        "b.item7 as M8,b.item8 as  M9,b.item9 as M10,B.item10 as M11,B.item11 as M12,b.item12 as M13,b.item13 as M14,b.item14 as M15,b.item15 as M16," & _
                        "b.item16 AS M17,b.item17 AS M18,b.item18 AS M19,b.item19 as M20,b.item20 AS M21,B.iteM21 as M22,B.ITEM22 AS M23,B.ITEM23 AS M24," & _
                        "b.item24 AS M25,b.item25 AS M26 ,b.item26 AS M27,b.item27 AS M28,b.item28 AS ITEM29,b.item29 as M30,b.item30 as M31,b.item31 AS M32," & _
                        "b.item32 AS M33, b.item33 AS M34 , b.item34 AS M35" & _
                        ",C.item0 as  C1,c.item1 as  C2,c.item2 as  C3,c.item3 as  C4,c.item4 as  C5,c.Item5 as  C6,c.item6 as  C7,c.item7 as  C8" & _
                        ",c.item8 as  C9,c.item9 as  C10,c.item10 as  C11,d.PicPrt" & _
                        " from sick A LEFT OUTER JOIN" & _
                        " chemic C ON A.ID = C.ID LEFT OUTER JOIN" & _
                        " micro B  ON A.ID = B.ID LEFT OUTER JOIN" & _
                        " micro_out D ON A.id = d.id and a.MICROID = substring(d.id,9,4)" & _
                        "where  a.chkdate>='" & Format(dtStart, "yyyy.MM.dd") & "' and a.chkdate<='" & Format(DateAdd("d", 1, dtEnd), "yyyy.MM.dd") & "'" & _
                        " order by A.chkdate,A.id "
                        
130         Set rsTmp = cnAccess.Execute(strSQL)
132         If rsTmp.EOF Then
134             WriteLog "Ave762A", "��131��", Format(dtStart, "yyyy.MM.dd") & "��" & Format(dtEnd, "yyyy.MM.dd") & ",�걾" & lngSampleNO & ",δ�ҵ����"
136             .Close
                Exit Function
            End If
        
138         ReDim aResults(0)
140         lngResultCount = 0
142         strBh = "***"
        
            'rsTmp.MoveFirst
144         Do Until rsTmp.EOF
146             If Val("" & rsTmp("id")) > 0 Then '�±걾
148                 strBh = "" & rsTmp("id")
150                 str���� = Replace("" & rsTmp("chkdate"), ".", "-") '����ʱ��
152                 str��� = ""
                
154                 For iField = 2 To rsTmp.Fields.count - 2
156                     If "" & rsTmp.Fields(iField).Value <> "" Then
158                         str��� = str��� & "|" & rsTmp.Fields(iField).Name & "|" & Replace(rsTmp.Fields(iField).Value, "mmol/l", "")
                        End If
                    Next
            
160                 If strBh <> "***" And str��� <> "" Then
162                     str�걾�� = Mid(strBh, 9)
164                     If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
166                     aResults(UBound(aResults)) = str����
168                     aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(str�걾��) '�������
170                     aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
172                     aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|��Һ"
174                     aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
176                     aResults(UBound(aResults)) = aResults(UBound(aResults)) & str���
178                     strPic = ""
180                     iCount = 0
182                     For i = 1 To 6
                            '---����JPG��ʽ��ͼ��
                            strImg = strPicPath & "\" & Mid(strBh, 1, 8) & "\" & Mid(strBh, 9) & "\H000" & i & ".jpg"
                            If Not gobjFSO.FileExists(strImg) Then strImg = strPicPath & "\" & Mid(strBh, 1, 8) & "\" & Mid(strBh, 9) & "\L000" & i & ".jpg"
                            If Not gobjFSO.FileExists(strImg) Then strImg = strPicPath & "\" & Mid(strBh, 1, 8) & "\" & Mid(strBh, 9) & "\F000" & i & ".jpg"
                            
                            '--- û������BMP��ʽ��ͼ��
184                         If Not gobjFSO.FileExists(strImg) Then strImg = strPicPath & "\" & Mid(strBh, 1, 8) & "\" & Mid(strBh, 9) & "\H000" & i & ".bmp"
186                         If Not gobjFSO.FileExists(strImg) Then strImg = strPicPath & "\" & Mid(strBh, 1, 8) & "\" & Mid(strBh, 9) & "\L000" & i & ".bmp"
                            If Not gobjFSO.FileExists(strImg) Then strImg = strPicPath & "\" & Mid(strBh, 1, 8) & "\" & Mid(strBh, 9) & "\F000" & i & ".bmp"

188                         If Dir(strImg) <> "" Then
                                Dim strImgFmt As String
                                strImgFmt = "03"
                                If UCase(Right(strImg, 3)) = "JPG" Then strImgFmt = "13"
190                             If intZIP = 0 Then
192                                 strPic = strPic & "^" & strImgTitle & "000" & i & ";1" & strImgFmt & ";0;" & strImg
                                Else
194                                 strPic = strPic & "^" & strImgTitle & "000" & i & ";2" & strImgFmt & ";0;" & strImg
                                End If
196                             iCount = iCount + 1
                            Else
198                             WriteLog "Ave762A", "ȱ��ͼ�Σ�" & strImg, ""
                            End If
                        Next

200                     If strPic <> "" Then
202                         If Left(strPic, 1) = "^" Then strPic = Mid(strPic, 2)
204                         aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & strPic
                        End If
                    End If
                End If
206             rsTmp.MoveNext
            Loop
        End With
    
208     ResultFromFile = aResults
        Exit Function
errH:
210 WriteLog "Ave762A", "��" & CStr(Erl()) & "��", Err.Description
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

    Dim intloop As Long, intCount As Integer
    Dim arrItems As Variant
    
    On Error GoTo ErrHand
    
    '�������ļ��ж�ȡ�������ļ��е�����cotf
    intCount = UBound(Split(strItems, ","))
    ReDim arrItems(intCount, 2)
    For intloop = 0 To intCount
        arrItems(intloop, 0) = Split(strItems, ",")(intloop)
    Next
    
    If objFilesys.FileExists(strIniFile) Then
        Set objStream = objFilesys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intloop = 0 To intCount
                    If Trim(arrLine(0)) = Trim(arrItems(intloop, 0)) Then
                        arrItems(intloop, 1) = Trim(arrLine(1))
                        'Left(tempusername,   InStr(1,   tempusername,   Chr(0))   -   1)
                        Exit For
                    End If
                Next
           End If
        Loop
        objStream.Close
        strReturn = ""
        For intloop = 0 To intCount
            strReturn = strReturn & "," & arrItems(intloop, 1)
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


