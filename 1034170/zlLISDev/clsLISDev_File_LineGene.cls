VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LineGene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: PCR ӫ�ⶨ�������� LineGene
'����:���ݲ���
'����:�¶�
'�û�:÷�ӿ�������ҽԺ
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
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strBh As String, strimg As String, strLastBH As String, strAllImg As String, strAllItems As String
    
    Dim cnErs As New ADODB.Connection
    Dim strFileErs As String, rsErs As ADODB.Recordset
    Dim strImgErs As String, str���� As String, str��Ŀ As String, str��� As String
    Dim strChar As String, str�ӿ����� As String
    Dim intMaxValue As Integer, i As Integer
    Dim strQC As String
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strChar = Replace(strFile, strFilePath & "\", "")
    strChar = Mid(strChar, 4, 1)
    
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
        
    ReDim aResults(0)
    
    strImgErs = ""
    strQC = "0"
    str�ӿ����� = Trim(ReadIni("�ӿ�", "����", App.Path & "\LineGane.ini"))
    If str�ӿ����� = "" Then
        WriteIni "�ӿ�", "����", "÷�ӿ�", App.Path & "\LineGane.ini"
        str�ӿ����� = "÷�ӿ�"
    End If
    
    If strFileErs <> "" Then
        strSQL = "SELECT sample_no  As ˳���, report_date as ����, " & _
                 "A.project_name as ��Ŀ,sample_name, A.con_Cal as ���, ct,ct_ref,con_min,result_con, " & _
                 "A.fluorescence_1, A.fluorescence_2, A.fluorescence_3, A.fluorescence_4, A.fluorescence_5," & _
                 "A.fluorescence_6, A.fluorescence_7, A.fluorescence_8, A.fluorescence_9, A.fluorescence_10," & _
                 "A.fluorescence_11, A.fluorescence_12, A.fluorescence_13, A.fluorescence_14, A.fluorescence_15," & _
                 "A.fluorescence_16, A.fluorescence_17, A.fluorescence_18, A.fluorescence_19, A.fluorescence_20," & _
                 "A.fluorescence_21, A.fluorescence_22, A.fluorescence_23, A.fluorescence_24, A.fluorescence_25," & _
                 "A.fluorescence_26, A.fluorescence_27, A.fluorescence_28, A.fluorescence_29, A.fluorescence_30" & _
                 " FROM qpcr_data AS A Where Format(A.report_date,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "' And Format(A.report_date,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & _
                 "' Order by sample_no" '& _
                'IIf(lngSampleNO = -1, "", " And val(Mid(sample_no,9))=" & lngSampleNO)
        Set rsErs = cnErs.Execute(strSQL)
        Do Until rsErs.EOF
            If Val("" & rsErs("˳���")) <> 0 Then
                str���� = Format(rsErs("����"), "yyyy-MM-dd")
                If strLastBH <> Trim("" & rsErs("˳���")) And strLastBH <> "" Then
                    If strAllItems <> "" Then
                        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        aResults(UBound(aResults)) = str����  '����ʱ��
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(strBh)   '�������
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "  '������
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "ѪҺ"  '�걾
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & strQC
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & strAllItems
                        If strAllImg <> "" Then aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & Mid(strAllImg, 2)
                    End If
                    strAllItems = ""
                    strAllImg = ""
                    
                ElseIf strLastBH = "" Then
                    strAllItems = ""
                    strAllImg = ""
                End If
                
                strBh = Trim("" & rsErs("˳���"))
                If Len(strBh) > 8 Then
                    strBh = Mid(strBh, 8)
                ElseIf Len(strBh) > 6 Then
                    strBh = Mid(strBh, 7)
                End If

                If str�ӿ����� = "÷�ӿ�" Then
                    strAllItems = strAllItems & "|" & rsErs!��Ŀ & "|" & E_Number("" & rsErs!���)
                    'str��Ŀ = Trim("" & rsErs!��Ŀ)
                Else
                    str��Ŀ = Trim("" & rsErs!sample_name)
                
                    If rsErs!result_con = 0 Then
                        'ȡCTֵ
                        If str��Ŀ <> "HBV" Then
                            str��� = Trim("" & rsErs!ct)
                            If str��� = "-1" Then
                                str��� = "�� (����)"
                            ElseIf Val(str���) > Val("" & rsErs!ct_ref) And Val("" & rsErs!ct_ref) > 0 Then
                                str��� = str��� & " (����)"
                            ElseIf str��� <> "" Then
                                str��� = str��� & " (����)"
                            End If
                        Else
                            str��Ŀ = ""
                        End If
                    Else
                        'HBV
                        If str��Ŀ = "HBV" Then
                            str��� = Trim("" & rsErs!���)
                            If Val(str���) < Val("" & rsErs!con_min) And Val("" & rsErs!con_min) <> 0 Then
                                str��� = "<" & E_Number(Val(rsErs!con_min))
                            Else
                                str��� = E_Number(Val(str���))
                            End If
                        Else
                            str��Ŀ = ""
                        End If
                    End If
                    
                    If str��Ŀ <> "" And str��� <> "" Then
                        strAllItems = strAllItems & "|" & str��Ŀ & "|" & str���
                    End If
                End If
                strImgErs = ""
                
                If strAllItems <> "" Then
                    For i = 6 To rsErs.Fields.count - 1
                        If rsErs.Fields(i).Value <> 0 Then
                            strImgErs = strImgErs & "," & Format((rsErs.Fields(i).Value / 6000) * 10, "0.000")
                        End If
                        If rsErs.Fields(i).Value > intMaxValue Then
                            intMaxValue = rsErs.Fields(i).Value
                        End If
                    Next
                    If strImgErs <> "" Then
                        strImgErs = "��;4;100,50,4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40|10-10,20-20,30-30,40-40,50-50,60-60,70-70,80-80,90-90,100-100;" & Mid(strImgErs, 2) & "~ӫ��ǿ��,12,100|ѭ����,55,10"
'                        If intMaxValue <= 6000 Then
'                            strImgErs = "����;3;100,50|0-0,10-10,20-20,30-30,40-40,50-50,60-60,70-70,80-80,90-90,100-100|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40;" & Mid(strImgErs, 2) & ";ӫ��ǿ��,5,100~ѭ����,55,1"
'                        ElseIf intMaxValue <= 10000 Then
'                            strImgErs = "����;3;100,50|0-0,10-10,20-20,30-30,40-40,50-50,60-60,70-70,80-80,90-90,100-100|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40;" & Mid(strImgErs, 2) & ";ӫ��ǿ��,5,100~ѭ����,55,1"
'                        End If
                        strAllImg = strAllImg & "^" & strImgErs
                    End If
                End If
                
                strLastBH = Trim("" & rsErs("˳���"))
            End If
            rsErs.MoveNext
        Loop
        
        
        If strAllItems <> "" Then
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            
            aResults(UBound(aResults)) = str����  '����ʱ��
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & strBh   '�������
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "  '������
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "ѪҺ"  '�걾
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & strAllItems
            If strAllImg <> "" Then aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & Mid(strAllImg, 2)
            strAllItems = ""
            strAllImg = ""
        End If
        
    End If
        

    ResultFromFile = aResults
End Function

Private Function E_Number(ByVal curIn As Double) As String
    '������תΪ��ѧ��������ʾ
    '����1000�Ĳ�ת,������λС��
    Dim strIn As String
    strIn = curIn
    
    If Len(Split(strIn, ".")(0)) > 0 Then
        strIn = Format(Val(strIn) / 10 ^ (Len(Split(strIn, ".")(0)) - 1), "0.00") & "E+" & Format(Len(Split(strIn, ".")(0)) - 1, "00")
    End If
    E_Number = IIf(Left(strIn, 1) = ".", "0" & strIn, strIn)
End Function
