VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XE2100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:XE-2100
'����:Sysmex
'����:�¶�
'�û�:����ʡ����ҽԺ
Private mstrCode As String

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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    Dim lng_Stx As Long, lng_Etx As Long, bln_QC As Boolean
    Dim strSampleNo As String, strSampleDate As String, int_Index As Integer
    Dim str_Code As String, str_Line As String, str_Item As String
    Dim strD1U As String, strTmp As String
    Dim str_Path As String, str_File  As String, str_imgName As String, str_RetuImage As String
    Dim bln_Rec As Boolean, strBacordID As String
    strCmd = Chr(&H6)
    mstrCode = mstrCode & strSource
    
    
        
    If Not mstrCode Like "*DI*D7G*" & Chr(&H3) & "*" Then Exit Sub
    lng_Stx = InStr(mstrCode, "DI")
    lng_Etx = InStr(lng_Stx, mstrCode, "D7G")
    lng_Etx = InStr(lng_Etx, mstrCode, Chr(&H3))
    
    strResult = ""
    
    Do While lng_Stx > 0 And lng_Etx > 0
        str_Code = Mid(mstrCode, lng_Stx + 1, lng_Etx - lng_Stx - 1)
        mstrCode = Mid(mstrCode, lng_Etx + 1)
        
        If Left(str_Code, 2) = "DI" Then
            bln_QC = False
        ElseIf ((Left(str_Code, 3) = "D1C") Or (Left(str_Code, 3) = "D2C")) Then
            bln_QC = True
        End If
        If bln_QC = False And Left(str_Code, 2) = "DI" Then
            strSampleNo = Trim(Mid(str_Code, 65, 15))
            If Len(strSampleNo) <= 0 Then Exit Do
            If Val(strSampleNo) <= 0 Then
                Do While Left(strSampleNo, 1) = "0"
                    strSampleNo = Mid(strSampleNo, 2)
                Loop
                bln_QC = True
            End If

            strSampleDate = Mid(str_Code, 43, 4)
            strSampleDate = strSampleDate & "-" & Mid(str_Code, 47, 2)
            strSampleDate = strSampleDate & "-" & Mid(str_Code, 49, 2)
            
            strSampleDate = strSampleDate & " " & Mid(str_Code, 51, 2)
            strSampleDate = strSampleDate & ":" & Mid(str_Code, 53, 2)
            strSampleDate = strSampleDate & ":" & Mid(str_Code, 55, 2)
            
            If IsDate(strSampleDate) Then
                strSampleDate = Format(CDate(strSampleDate), "yyyy-MM-dd HH:mm:ss")
            Else
                strSampleDate = Format(Now, "yyyy-MM-dd HH:mm:ss")
            End If
            strD1U = Mid(str_Code, InStr(str_Code, "D1U"), 132)
            str_Line = Mid(str_Code, InStr(str_Code, "D2U"), 216)
            str_Item = ""
            str_RetuImage = ""
            
            If Len(strD1U) >= 132 Then
                strTmp = Trim(GetValues(Mid(strD1U, 93, 7), 3))
                If strTmp <> "" Then str_Item = str_Item & "|WBC-BF|" & strTmp
                
                strTmp = Trim(GetValues(Mid(strD1U, 100, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|RBC-BF|" & strTmp
                
                strTmp = Trim(GetValues(Mid(strD1U, 113, 7), 3))
                If strTmp <> "" Then str_Item = str_Item & "|MN#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(strD1U, 120, 5), 1))
                If strTmp <> "" Then str_Item = str_Item & "|MN%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(strD1U, 125, 7), 3))
                If strTmp <> "" Then str_Item = str_Item & "|PMN#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(strD1U, 132, 5), 1))
                If strTmp <> "" Then str_Item = str_Item & "|PMN%|" & strTmp
            End If
            If Len(str_Line) >= 216 Then
                
                strTmp = Trim(GetValues(Mid(str_Line, 11, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|WBC|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 17, 5), 3))
                If strTmp <> "" Then str_Item = str_Item & "|RBC|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 22, 5), 1))
                If strTmp <> "" Then str_Item = str_Item & "|HGB|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 27, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|HCT|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 32, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|MCV|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 37, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|MCH|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 42, 5), 1))
                If strTmp <> "" Then str_Item = str_Item & "|MCHC|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 47, 5), 1))
                If strTmp <> "" Then str_Item = str_Item & "|PLT|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 52, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|LYMPH%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 57, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|MONO%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 62, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|NEUT%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 67, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|EO%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 72, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|BASO%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 77, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|LYMPH#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 83, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|MONO#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 89, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|NEUT#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 95, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|EO#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 101, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|BASO#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 107, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|RDW-CV|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 112, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|RDW-SD|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 117, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|PDW|" & strTmp
                 
                strTmp = Trim(GetValues(Mid(str_Line, 122, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|MPV|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 127, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|P-LCR|" & strTmp
                
                bln_Rec = Not (Val(Mid(str_Line, 132, 5)) / 10 ^ 3 = 0)
                
                strTmp = Trim(GetValues(Mid(str_Line, 132, 5), 3))
                If strTmp <> "" Then str_Item = str_Item & "|RET%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 137, 5), 5))
                If strTmp <> "" Then str_Item = str_Item & "|RET#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 142, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|IRF|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 147, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|LFR|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 152, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|MFR|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 157, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|HFR|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 162, 5), 3))
                If strTmp <> "" Then str_Item = str_Item & "|PCT|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 167, 6), 1))
                If strTmp <> "" Then str_Item = str_Item & "|NRBC%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 173, 6), 2))
                If strTmp <> "" Then str_Item = str_Item & "|NRBC#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 179, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|IG#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 185, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|IG%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 190, 6), 2))
                If strTmp <> "" Then str_Item = str_Item & "|HPC#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 196, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|RET-He|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 201, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|IPF|" & strTmp
                
                
            End If
            
            
            If strSampleNo <> "" And str_Item <> "" Then
            
                If bln_QC Then
                    strResult = strResult & "||" & strSampleDate & "|" & strSampleNo & "| |Ѫ��|1" & str_Item
                Else
                    strSampleNo = Val(strSampleNo)
                    If Len(strSampleNo) >= 12 Then
                        Call ReadConf(App.Path & "\XE2100.ini", "B_ID", strBacordID)
                        If strBacordID = "" Then
                            '�������⴦���������3λ��Ϊҽ��ID�������걾�ŷ��ء�
                            If Val(Mid(strSampleNo, 3)) <> 0 Then
                                strResult = strResult & "||" & strSampleDate & "|" & Val(Mid(strSampleNo, 3)) & "^^" & strSampleNo & "| |Ѫ��|0" & str_Item
                            Else
                                strResult = strResult & "||" & strSampleDate & "|" & "^^" & strSampleNo & "| |Ѫ��|0" & str_Item
                            End If
                        Else
                            strResult = strResult & "||" & strSampleDate & "|" & "^^" & strSampleNo & "| |Ѫ��|0" & str_Item
                        End If
                    Else
                        strResult = strResult & "||" & strSampleDate & "|" & strSampleNo & "| |Ѫ��|0" & str_Item
                    End If
                End If
            Else
                Exit Sub
            End If
            
            'ͼ������
            str_Path = App.Path & "\Gram"
            If Dir(str_Path, vbDirectory) = "" Then
                MkDir str_Path
            End If
            If InStr(str_Code, "D3U") > 0 Then
                str_imgName = "RBC"
                str_File = Mid(Replace(strSampleDate, "-", "_"), 1, 10) & "_" & Val(strSampleNo) & "_" & str_imgName & ".gif"
                
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D3U"))
            End If
            If InStr(str_Code, "D4U") > 0 Then
                str_imgName = "PLT"
                str_File = Mid(Replace(strSampleDate, "-", "_"), 1, 10) & "_" & Val(strSampleNo) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D4U"))
            End If
            If InStr(str_Code, "D1G") > 0 Then
                str_imgName = "SDIFF"
                str_File = Mid(Replace(strSampleDate, "-", "_"), 1, 10) & "_" & Val(strSampleNo) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D1G"))
            End If
            If InStr(str_Code, "D2G") > 0 Then
                str_imgName = "SBASO"
                str_File = Mid(Replace(strSampleDate, "-", "_"), 1, 10) & "_" & Val(strSampleNo) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D2G"))
            End If

            If str_RetuImage <> "" Then strResult = strResult & vbCrLf & Mid(str_RetuImage, 2)
            
        Else
            'û���ݣ���ʱ������
        End If
        
        If mstrCode Like "*DI*D7G*" & Chr(&H3) & "*" Then
            lng_Stx = InStr(mstrCode, "DI")
            lng_Etx = InStr(lng_Stx, mstrCode, "D7G")
            lng_Etx = InStr(lng_Etx, mstrCode, Chr(&H3))
        Else
             lng_Stx = 0
             lng_Etx = 0
        End If

    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
End Function

Private Function Save_image(ByVal str_Img As String, ByVal str_FileName As String, ByVal str_Data As String, ByVal lng_Start As Long) As String
    Dim lng_Lens As Long, str_ImgData As String, T_ImgInfo As typHistGramInfo
    
    If lng_Start <= 0 Then Exit Function
    
    If str_Img = "RBC" Or str_Img = "PLT" Or str_Img = "WBC DISCRI" Then
        lng_Lens = Val(Mid(str_Data, lng_Start + 22, 6)) - 12
        If lng_Lens <= 0 Then Exit Function
        str_ImgData = Mid(str_Data, lng_Start + 41, lng_Lens)
        T_ImgInfo.lnglower = Val(Mid(str_Data, lng_Start + 29, 4))
        T_ImgInfo.lngupper = Val(Mid(str_Data, lng_Start + 33, 4))
        T_ImgInfo.lngmaxx = Val(Mid(str_Data, lng_Start + 16, 3))
        T_ImgInfo.lngmaxy = Val(Mid(str_Data, lng_Start + 19, 3))
        T_ImgInfo.lngresver1 = 0: T_ImgInfo.lngresver2 = 0
        T_ImgInfo.lngstoppos = IIf(str_Img = "PLT", 40, 46)
        
        If T_ImgInfo.lngmaxx = 0 Or T_ImgInfo.lngmaxy = 0 Then
            WriteLog "saveImage", "�������ݴ���", "x" & T_ImgInfo.lngmaxx & ",y" & T_ImgInfo.lngmaxy
            Exit Function
        End If
        If uf_xehist(str_ImgData, lng_Lens, T_ImgInfo, "c:\temp1.bmp", str_FileName) = 1 Then
             Save_image = "^" & str_Img & ";105;1;" & str_FileName
        End If
    Else
        lng_Lens = Val(Mid(str_Data, lng_Start + 22, 6)) - 1
        If lng_Lens <= 0 Then Exit Function
        str_ImgData = Mid(str_Data, lng_Start + 29, lng_Lens)
        If uf_xescat(str_ImgData, lng_Lens, "c:\Temp2.bmp", str_FileName) = 1 Then
            Save_image = "^" & str_Img & ";104;1;" & str_FileName
        End If
    End If
    
End Function

Private Function GetValues(ByVal str_in As String, int_i As Integer) As String
    '����:��ʽ������ֵ
    'str_in:����ʽ�����ַ���
    'int_i:�����������ʱ,������С��λ��
    
    If str_in = "*0000" Then
        GetValues = "----"
    ElseIf str_in = Space(5) Or str_in = Space(6) Or str_in = Space(7) Then
        GetValues = " "
    Else
        GetValues = Val(str_in) / 10 ^ int_i
    End If
    
End Function


Private Sub ReadConf(strIniFile As String, ByVal strItems As String, ByRef strReturn As String)
    '�������ļ�
    'strIniFile :ini�����ļ���
    'strItems   :Ҫ��ȡ���ֶΣ��磺DATABASE NAME=th_xlb�е� DATABASE NAME,����ֶ���,�ָ�
    'strReturn  :��ȡ�ķ���ֵ����:DATABASE NAME=th_xlb�е� th_xlb ���������ֵ��,�ָ�
    
    Dim arrLine
    Dim strLine As String

    Dim blnOpen As Boolean
    Dim objFileSys As New FileSystemObject
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
    
    If objFileSys.FileExists(strIniFile) Then
        Set objStream = objFileSys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intloop = 0 To intCount
                    If Trim(arrLine(0)) = arrItems(intloop, 0) Then
                        arrItems(intloop, 1) = arrLine(1)
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

