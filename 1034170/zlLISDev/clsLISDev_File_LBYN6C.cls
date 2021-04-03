VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LBYN6C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Ѫ���� ������ LBY-N6C
'����:������
'����:�¶�
'�û�:����ҽԺ,�Ͼ���������ҽԺ,��������(LBY-N6B)
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
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean, strImg As String, iField As Integer, i As Integer
    
    Dim cnErs As New ADODB.Connection
    Dim strFileErs As String, rsErs As ADODB.Recordset
    Dim strImgErs As String, str���� As String
    Dim strChar As String
    Dim intMaxValue As Integer
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strChar = Replace(strFile, strFilePath & "\", "")
    strChar = Mid(strChar, 4, 1)
    
    strFile = strFilePath & "\DN6" & strChar & Format(dtStart, "YYMM") & ".MDB"
    strFileErs = strFilePath & "\DEsr" & Format(dtStart, "YYMM") & ".MDB"
    
    If Dir(strFile) = "" Then Exit Function
    If Trim(Dir(strFileErs)) <> "" Then 'Ѫ������
        With cnErs
        .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then strFileErs = ""
        End With
    Else
        strFileErs = ""
    End If
    
    With cnAccess
        
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strSQL = "Select val(Mid(TestWholeno,9)) As ˳���, Patientname As ����, PatientSex As �Ա�, Testdate As ����, Patientanalysisresult As ���, Qxdata1 As ����ֵ," & vbNewLine & _
                "       Qxdata5 As ����ֵ, Qxdata10 As ����ֵ, Xjdata As ճ��, Rltcalcparam1 As ѹ��, Rltcalcparam2 As Ѫ��," & vbNewLine & _
                "       Rltcalcdata1 As ��ԭճ�ȵ���ֵ, Rltcalcdata2 As ��ԭճ������ֵ, Rltcalcdata3 As ��ԭճ�ȸ���ֵ," & vbNewLine & _
                "       Rltcalcdata4 As Ѫ������kֵ, Rltcalcdata5 As �ۼ�ָ��, Rltcalcdata6 As ����ָ��, Rltcalcdata7 As ����ָ��," & vbNewLine & _
                "       Rltcalcdata8 As ��Ӿָ��, Qxstrparam1 As ����ʱ��, Qxstrparam5 As ����ʱ��, Qxstrparam10 As ����ʱ��" & vbNewLine & _
                "From N6" & strChar & "data Where Format(Testdate,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "' And Format(Testdate,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "' " & _
                IIf(lngSampleNO = -1, "", " And val(Mid(TestWholeno,9))=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        Do While Not rsTmp.EOF
            If strBh <> rsTmp("˳���") Then '�±걾
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(rsTmp("����"), "yyyy-mm-dd hh:mm:ss")  '����ʱ��
                
                str���� = Format(rsTmp("����"), "yyyy-mm-dd")
                
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val("" & rsTmp("˳���"))   '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      " "  '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      "ѪҺ"  '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|0" '& _
                   ' Val(rsTmp("�ʿ�")) '�ʿ�
                    
                strBh = rsTmp("˳���")
            End If
            
            For iField = 5 To 19
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    iField - 4 & "|" & Val(IIf(IsNull(rsTmp.Fields(iField).Value), "0", rsTmp.Fields(iField).Value)) '���
            Next
            '-- Ѫ�����ͼ������
            '-- ͼ������X-�ᣬY����
            '-- ��ʽ������;��������;���߼��������
            '        ����  �������ݣ�Y����,X����|X����-X������ʾ������,....|Y����-Y������ʾ������,....
            '              ���߼��������:����1��Y����1,����1��Y����2.....|����2��Y����1,����2��Y����2,...~����Y�ӱ�1,����Y�ӱ�2
            strImg = vbCrLf & "ճ����������;2;20,200|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20;"
            
            If Trim("" & rsTmp("�Ա�")) = "Ů" Then
'                ND0 = 9.25: QB0 = 10: ND1 = 4.4: QB1 = 150
'                ND2 = 6.5: QB2 = 10: ND3 = 3.65: QB3 = 150
                strImg = strImg & "9.25,10,4.4,150|6.5,10,3.65,150"
            Else
'                ND0 = 9.58: QB0 = 10: ND1 = 4.6: QB1 = 150
'                ND2 = 6.8: QB2 = 10: ND3 = 3.73: QB3 = 150
                strImg = strImg & "9.58,10,4.6,150|6.8,10,3.73,150"
            End If
            strImg = strImg & "~10-" & rsTmp!����ֵ & ",60-" & rsTmp!����ֵ & ",150-" & rsTmp!����ֵ  '���
            strImg = strImg & ";VIS(mPa.s),15,20~SHR(1/S),195,1" '������
            
            '-- Ѫ����ͼ������
            strImgErs = ""
            If strFileErs <> "" Then
                strSQL = "SELECT A.SerialNo, A.PatientSample, A.TestDataIndex, A.TestDate, A.TestData1, A.TestData2, A.TestData3, A.TestData4, A.TestData5, A.TestData6, A.TestData7, A.TestData8, A.TestData9, A.TestData10, A.TestData11, A.TestData12, A.TestData13, A.TestData14, A.TestData15, A.TestData16, A.TestData17, A.TestData18, A.TestData19, A.TestData20, A.TestData21, A.TestData22, A.TestData23, A.TestData24, A.TestData25, A.TestData26, A.TestData27, A.TestData28, A.TestData29, A.TestData30" & _
                        " FROM EsrData AS A Where Format(A.Testdate,'yyyy-MM-dd')>='" & Format(rsTmp("����"), "yyyy-mm-dd") & "' And Format(A.Testdate,'yyyy-MM-dd')<='" & Format(rsTmp("����"), "yyyy-mm-dd") & _
                        "' And val(TestDataIndex) >0 And A.SerialNo=" & Val(rsTmp("˳���"))
                Set rsErs = cnErs.Execute(strSQL)
                Do Until rsErs.EOF
                    For i = 4 To rsErs.Fields.count - 1
                        strImgErs = strImgErs & "," & rsErs.Fields(i).Value
                        If rsErs.Fields(i).Value > intMaxValue Then
                            intMaxValue = rsErs.Fields(i).Value
                        End If
                    Next
                    rsErs.MoveNext
                Loop
                If strImgErs <> "" Then
                    If intMaxValue <= 36 Then
                        strImgErs = "Ѫ������;3;36,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36;" & Mid(strImgErs, 2) & ";Ѫ��ֵ(mm),5,36~ʱ��(m),55,1"
                    ElseIf intMaxValue <= 40 Then
                        strImgErs = "Ѫ������;3;40,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40;" & Mid(strImgErs, 2) & ";Ѫ��ֵ(mm),5,40~ʱ��(m),55,1"
                    ElseIf intMaxValue <= 44 Then
                        strImgErs = "Ѫ������;3;44,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40,44-44;" & Mid(strImgErs, 2) & ";Ѫ��ֵ(mm),5,44~ʱ��(m),55,1"
                    ElseIf intMaxValue <= 48 Then
                        strImgErs = "Ѫ������;3;48,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40,44-44,48-48;" & Mid(strImgErs, 2) & ";Ѫ��ֵ(mm),5,48~ʱ��(m),55,1"
                    ElseIf intMaxValue <= 52 Then
                        strImgErs = "Ѫ������;3;52,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40,44-44,48-48,52-52;" & Mid(strImgErs, 2) & ";Ѫ��ֵ(mm),5,52~ʱ��(m),55,1"
                    ElseIf intMaxValue <= 56 Then
                        strImgErs = "Ѫ������;3;56,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40,44-44,48-48,52-52,56-56;" & Mid(strImgErs, 2) & ";Ѫ��ֵ(mm),5,56~ʱ��(m),55,1"
                    ElseIf intMaxValue <= 60 Then
                        strImgErs = "Ѫ������;3;60,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40,44-44,48-48,52-52,56-56,60-60;" & Mid(strImgErs, 2) & ";Ѫ��ֵ(mm),5,36~ʱ��(m),55,1"
                    End If
'                    strImgErs = "Ѫ������;3;60,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40,44-44,48-48,52-52,56-56,60-60;" & Mid(strImgErs, 2) & ";Ѫ��ֵ(mm),5,60~ʱ��(m),55,1"
                End If
            End If
            
            If strImgErs <> "" Then strImg = strImg & "^" & strImgErs
            
            aResults(lngResultCount) = aResults(lngResultCount) & strImg
            
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function

