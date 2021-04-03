VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_XE5000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:����ˮ�ߵ� XE-5000
'����:ϣɭ����(SysMex)
'����:�¶�
'�û�:��ú����ҽԺ

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
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream, strLine As String

    Dim str���� As String, str��Ŀ As String, str�걾�� As String, strImg As String
    Dim strFilePath As String, strFilename As String, strImgPath As String, strImgFilePath As String
    Dim arrLine
    Call ReadConf(App.Path & "\XE5000.ini", "IMG_PATH", strImgPath)
    
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    strFilename = Dir(strFilePath & "\*.CDF")
    ReDim aResults(0)
    
    Do While strFilename <> ""
        
        If objFile.FileExists(strFilePath & "\" & strFilename) Then
            Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFilename)
            
            Do While Not objStream.AtEndOfStream
                strLine = UCase(objStream.ReadLine)
                arrLine = Split(strLine, ",")
                If arrLine(0) = "0" Then
                    str��Ŀ = "": strImg = ""
                    str���� = Split(strLine, ",")(1) & " " & Split(strLine, ",")(2)
                    If CDate(str����) >= dtStart And CDate(str����) <= dtEnd Then
                        str�걾�� = Split(strLine, ",")(3)
                        If lngSampleNO <> -1 Then
                            If lngSampleNO <> Val(str�걾��) Then
                                str���� = "": str�걾�� = ""
                                Exit Do
                            End If
                        End If
                    Else
                        str���� = "": str�걾�� = ""
                        Exit Do
                    End If
                    
                ElseIf arrLine(0) = "1" Then
                    '������
                    str��Ŀ = str��Ŀ & "|" & Split(strLine, ",")(1) & "|" & Split(strLine, ",")(3)
                ElseIf arrLine(0) = "3" Then
                    '�滻ͼ��·��
                    strImgFilePath = Split(strLine, ",")(3)
                    If InStr(strImgPath, ",") > 0 Then
                        strImgFilePath = Replace(strImgFilePath, Trim(Split(strImgPath, ",")(0)), Trim(Split(strImgPath, ",")(1)))
                    End If
                    If InStr("HRBC,HPLT", Split(strLine, ",")(2)) > 0 Then
                        strImg = strImg & "^" & Split(strLine, ",")(2) & ";105;0;" & strImgFilePath
                    Else
                        strImg = strImg & "^" & Split(strLine, ",")(2) & ";104;0;" & strImgFilePath
                    End If
                ElseIf arrLine(0) = "9" Then
                    Exit Do
                End If
            Loop
            
            If str��Ŀ <> "" And str���� <> "" And str�걾�� <> "" Then
                str��Ŀ = Mid(str��Ŀ, 2)
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                If Len(str�걾��) >= 12 Then str�걾�� = "^^" & str�걾��
                aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |Ѫ��|0|" & str��Ŀ
                If strImg <> "" Then
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & Mid(strImg, 2)
                End If
                
                str��Ŀ = "": str���� = "": str�걾�� = "": strImg = ""
            End If
            
            objStream.Close
            Set objStream = Nothing
            
            '���˾Ͱ��ļ��ŵ�BAK��
            If Not objFile.FolderExists(strFilePath & "\Bak") Then Call objFile.CreateFolder(strFilePath & "\Bak")
            objFile.CopyFile strFilePath & "\" & strFilename, strFilePath & "\Bak\" & strFilename, True
            objFile.DeleteFile strFilePath & "\" & strFilename, True
            
        End If
        
        strFilename = Dir
    Loop
    ResultFromFile = aResults
End Function

Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    
    'iSendStep����ǰ���Ͳ����Ĳ��衣�����������ͽ������򷵻�0�������Զ���1���ء�
    'strResponse����ǰ��������Ӧ����
    'strSampleInfo�����͵ı걾��Ϣ��
    '   �걾֮����||�ָ�
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ���
    '   ��5�����أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��

    Dim var_�걾 As Variant, i_Count As Integer, i_Loop As Integer, str_��¼ As String, str_�걾�� As String
    Dim i_���� As Integer, str_���� As String, str_������Ŀ As String, str_�걾���� As String
    Dim str_SendItem As String, str_EndString As String, str_Tmp As String, lngRow As Long
    Dim str���� As String, arr����(3) As String, strCDF As String
    
    If strSampleInfo <> "" Then WriteLog "˫��ͨѶ(Xe5000 File Begin)��", strSampleInfo, ""
    var_�걾 = Split(strSampleInfo, "||")
    SendSample = ""
   
    Call ReadConf(App.Path & "\XE5000.ini", "SAVECDF_PATH,RET_PATH", str����)
    If UBound(Split(str����, ",")) = 1 Then
        arr����(0) = Split(str����, ",")(0)
        arr����(1) = Split(str����, ",")(1)

    Else
        arr����(0) = App.Path
        arr����(1) = App.Path
    End If
    For i_Count = LBound(var_�걾) To UBound(var_�걾)
        str_��¼ = var_�걾(i_Count)
        
        If Val(Split(str_��¼, "|")(1)) = 0 Then
            str_�걾�� = Trim(Split(str_��¼, "|")(5))
        Else
            str_�걾�� = Val(Split(str_��¼, "|")(1))
        End If
        str_�걾���� = Split(str_��¼, "|")(3)
        i_���� = Val(Split(str_��¼, "|")(4))
        str_���� = Trim(Split(str_��¼, "|")(5))
        '�̶�
        str_������Ŀ = ""
        
        For i_Loop = 10 To UBound(Split(str_��¼, "|"))
            If Trim(Split(str_��¼, "|")(i_Loop)) <> "" Then
                str_Tmp = Trim(Split(str_��¼, "|")(i_Loop))
                Call ReadConf(App.Path & "\XE5000.ini", str_Tmp, str����)
                If str���� <> "" Then
                    str_������Ŀ = str_������Ŀ & Trim(str����) & "+"
                End If
            End If
        Next
        
        If str_������Ŀ <> "" And str_���� <> "" Then
            str_������Ŀ = Mid(str_������Ŀ, 1, Len(str_������Ŀ) - 1)
            SendSample = SendSample & str_���� & "," & str_������Ŀ & "|"
            
           '�걾����(19)*,��������(20)*,�걾��(20)*,�걾����,�ѱ�,�ͼ����,�ͼ�ҽ��,����ҽ��,������*,��������*,
           '�Ա� (2),��������,����,����,���䵥λ,�����,�ͼ�����,��������,��ӡ��־,�����־,
           '�����־,���,��ע,�����ֶ�,����ƴ����,����ʱ��/��������,����,������
           strCDF = Split(str_��¼, "|")(0) & ",XE-5000," & str_���� & ",,,,,," & str_���� & "," & str_�걾�� & ","
           strCDF = strCDF & ",,,,,,,,,," & ",,,,,,," & str_����
           str_Tmp = Trim(arr����(0) & "\" & Format(Now, "yyyyMMdd") & str_���� & Format(Now, "HHmmss") & ".CDF")
           Call WritTxt(str_Tmp, strCDF)
        End If
        
    Next
    If SendSample <> "" Then
        str_Tmp = Trim(arr����(1) & "\RET.TXT")
        Call WritTxt(str_Tmp, SendSample)
        SendSample = ""
    End If
        
End Function

Private Function WritTxt(ByVal strFilename As String, strInfo As String)
    Dim objFileSystem As New FileSystemObject
    Dim objStream As TextStream
    
    If Not objFileSystem.FileExists(strFilename) Then Call objFileSystem.CreateTextFile(strFilename)
    Set objStream = objFileSystem.OpenTextFile(strFilename, ForAppending)
    objStream.WriteLine (strInfo)
    objStream.Close
    Set objStream = Nothing
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
    Dim str_�걾�� As String, str_���� As String, int_Index As Integer
    Dim str_Code As String, str_Line As String, str_Item As String
    
    Dim str_Path As String, str_File  As String, str_imgName As String, str_RetuImage As String
    Dim bln_Rec As Boolean
    strCmd = Chr(&H6)
    strReserved = strSource
    If Not strReserved Like "*DI*" & Chr(&H3) & "DR*" & Chr(&H3) & "*" Then Exit Sub
    lng_Stx = InStr(strReserved, "DI")
    lng_Etx = InStr(lng_Stx, strReserved, Chr(&H3))
    
    str_Code = Mid(strReserved, lng_Stx + 1, lng_Etx - lng_Stx - 1)
    
    lng_Stx = InStr(strReserved, "DR")
    lng_Etx = InStr(lng_Stx, strReserved, Chr(&H3))
    
    str_Code = str_Code & Mid(strReserved, lng_Stx + 1, lng_Etx - lng_Stx - 1)
    
    strReserved = Mid(strReserved, lng_Etx + 1)
    strResult = ""
    If Left(str_Code, 2) = "DI" Then
        bln_QC = False
    ElseIf ((Left(str_Code, 3) = "D1C") Or (Left(str_Code, 3) = "D2C")) Then
        bln_QC = True
    End If
    If bln_QC = False And Left(str_Code, 2) = "DI" Then
        str_�걾�� = Mid(str_Code, 65, 15)
        If Val(str_�걾��) <= 0 Then Exit Sub
        
        str_���� = Mid(str_Code, 43, 4)
        str_���� = str_���� & "-" & Mid(str_Code, 47, 2)
        str_���� = str_���� & "-" & Mid(str_Code, 49, 2)
        If IsDate(str_����) Then
            str_���� = Format(str_����, "yyyy-MM-dd")
        Else
            str_���� = Format(Now, "yyyy-MM-dd")
        End If
        
        str_Line = Mid(str_Code, InStr(str_Code, "D2U"), 216)
        str_Line = str_Line & Mid(str_Code, InStr(str_Code, "D6U"), 255)
        
        If Len(str_Line) >= 471 Then
            str_Item = "|WBC|" & Val(Mid(str_Line, 11, 6)) / 10 ^ 3
            str_Item = str_Item & "|RBC|" & GetValues(Mid(str_Line, 17, 5), 3)
            str_Item = str_Item & "|HGB|" & GetValues(Mid(str_Line, 22, 5), 1)
            str_Item = str_Item & "|HCT|" & GetValues(Mid(str_Line, 27, 5), 2)
            str_Item = str_Item & "|MCV|" & GetValues(Mid(str_Line, 32, 5), 2)
            str_Item = str_Item & "|MCH|" & GetValues(Mid(str_Line, 37, 5), 2)
            str_Item = str_Item & "|MCHC|" & GetValues(Mid(str_Line, 42, 5), 1)
            str_Item = str_Item & "|PLT|" & GetValues(Mid(str_Line, 47, 5), 1)
            str_Item = str_Item & "|LYMPH%|" & GetValues(Mid(str_Line, 52, 5), 2)
            str_Item = str_Item & "|MONO%|" & GetValues(Mid(str_Line, 57, 5), 2)
            str_Item = str_Item & "|NEUT%|" & GetValues(Mid(str_Line, 62, 5), 2)
            str_Item = str_Item & "|EO%|" & GetValues(Mid(str_Line, 67, 5), 2)
            str_Item = str_Item & "|BASO%|" & GetValues(Mid(str_Line, 72, 5), 2)
            str_Item = str_Item & "|LYMPH#|" & GetValues(Mid(str_Line, 77, 6), 3)
            str_Item = str_Item & "|MONO#|" & GetValues(Mid(str_Line, 83, 6), 3)
            str_Item = str_Item & "|NEUT#|" & GetValues(Mid(str_Line, 89, 6), 3)
            str_Item = str_Item & "|EO#|" & GetValues(Mid(str_Line, 95, 6), 3)
            str_Item = str_Item & "|BASO#|" & GetValues(Mid(str_Line, 101, 6), 3)
            str_Item = str_Item & "|RDW-CV|" & GetValues(Mid(str_Line, 107, 5), 2)
            str_Item = str_Item & "|RDW-SD|" & GetValues(Mid(str_Line, 112, 5), 2)
            str_Item = str_Item & "|PDW|" & GetValues(Mid(str_Line, 117, 5), 2)
            str_Item = str_Item & "|MPV|" & GetValues(Mid(str_Line, 122, 5), 2)
            str_Item = str_Item & "|P-LCR|" & GetValues(Mid(str_Line, 127, 5), 2)
            
            bln_Rec = Not (Val(Mid(str_Line, 132, 5)) / 10 ^ 3 = 0)
            str_Item = str_Item & "|RET%|" & GetValues(Mid(str_Line, 132, 5), 3)
            str_Item = str_Item & "|RET#|" & GetValues(Mid(str_Line, 137, 5), 5)
            str_Item = str_Item & "|IRF|" & GetValues(Mid(str_Line, 142, 5), 2)
            str_Item = str_Item & "|LFR|" & GetValues(Mid(str_Line, 147, 5), 2)
            str_Item = str_Item & "|MFR|" & GetValues(Mid(str_Line, 152, 5), 2)
            str_Item = str_Item & "|HFR|" & GetValues(Mid(str_Line, 157, 5), 2)
            
            str_Item = str_Item & "|PCT|" & GetValues(Mid(str_Line, 162, 5), 3)
            str_Item = str_Item & "|NRBC%|" & GetValues(Mid(str_Line, 167, 6), 1)
            str_Item = str_Item & "|NRBC#|" & GetValues(Mid(str_Line, 173, 6), 2)
            
            str_Item = str_Item & "|IG#|" & GetValues(Mid(str_Line, 179, 6), 3)
            str_Item = str_Item & "|IG%|" & GetValues(Mid(str_Line, 185, 5), 2)
            str_Item = str_Item & "|HPC#|" & GetValues(Mid(str_Line, 190, 6), 2)
            str_Item = str_Item & "|RET-He|" & GetValues(Mid(str_Line, 196, 5), 2)
            str_Item = str_Item & "|IPF|" & GetValues(Mid(str_Line, 201, 5), 2)
            
            str_Item = str_Item & "|WBC-B|" & GetValues(Mid(str_Line, 228, 5), 3)
            str_Item = str_Item & "|WBC-D|" & GetValues(Mid(str_Line, 234, 5), 3)
            str_Item = str_Item & "|NEUT#&|" & GetValues(Mid(str_Line, 240, 5), 3)
            str_Item = str_Item & "|NEUT%&|" & GetValues(Mid(str_Line, 246, 4), 2)
            str_Item = str_Item & "|LYMP#&|" & GetValues(Mid(str_Line, 251, 5), 3)
            str_Item = str_Item & "|LYMP%&|" & GetValues(Mid(str_Line, 257, 4), 2)
            str_Item = str_Item & "|HFLC#|" & GetValues(Mid(str_Line, 262, 5), 3)
            str_Item = str_Item & "|HFLC%|" & GetValues(Mid(str_Line, 268, 4), 2)
            str_Item = str_Item & "|Area#|" & GetValues(Mid(str_Line, 273, 5), 3)
            str_Item = str_Item & "|Area%|" & GetValues(Mid(str_Line, 279, 4), 2)
            str_Item = str_Item & "|NRBC+W|" & GetValues(Mid(str_Line, 284, 5), 3)
            str_Item = str_Item & "|NEUT-X|" & GetValues(Mid(str_Line, 289, 5), 2)
            str_Item = str_Item & "|RBC-O|" & GetValues(Mid(str_Line, 295, 4), 3)
            str_Item = str_Item & "|PLT-O|" & GetValues(Mid(str_Line, 300, 4), 1)
            str_Item = str_Item & "|PLT-I|" & GetValues(Mid(str_Line, 305, 4), 1)
            str_Item = str_Item & "|HF-BF#|" & GetValues(Mid(str_Line, 310, 6), 1)
            str_Item = str_Item & "|HF-BFB|" & GetValues(Mid(str_Line, 317, 4), 2)
            str_Item = str_Item & "|TC-BF#|" & GetValues(Mid(str_Line, 322, 6), 1)
            str_Item = str_Item & "|FRC#|" & GetValues(Mid(str_Line, 340, 4), 4)
            str_Item = str_Item & "|FRC%|" & GetValues(Mid(str_Line, 345, 4), 3)
            str_Item = str_Item & "|RBC-He|" & GetValues(Mid(str_Line, 350, 4), 2)
            str_Item = str_Item & "|Delta-He|" & GetValues(Mid(str_Line, 355, 5), 2)
            str_Item = str_Item & "|RET-Y|" & GetValues(Mid(str_Line, 360, 5), 2)
            str_Item = str_Item & "|RBC-Y|" & GetValues(Mid(str_Line, 365, 5), 2)
            str_Item = str_Item & "|IRF-Y|" & GetValues(Mid(str_Line, 370, 5), 2)
            str_Item = str_Item & "|RPI|" & GetValues(Mid(str_Line, 376, 4), 1)
            str_Item = str_Item & "|LScRBC|" & GetValues(Mid(str_Line, 381, 4), 2)
            str_Item = str_Item & "|HScRBC|" & GetValues(Mid(str_Line, 386, 4), 2)
            str_Item = str_Item & "|MicroRBC|" & GetValues(Mid(str_Line, 391, 4), 2)
            str_Item = str_Item & "|MacroRBC|" & GetValues(Mid(str_Line, 396, 4), 2)
            str_Item = str_Item & "|H-IPF|" & GetValues(Mid(str_Line, 401, 4), 2)
            str_Item = str_Item & "|IPF#|" & GetValues(Mid(str_Line, 406, 4), 2)
            str_Item = str_Item & "|PLT-X|" & GetValues(Mid(str_Line, 411, 4), 2)
            
        End If
        
        
        If Val(str_�걾��) > 0 And str_Item <> "" Then
            str_�걾�� = Val(str_�걾��)
            strResult = str_���� & "|" & IIf((Len(str_�걾��) >= 12), "^^", "") & str_�걾�� & "| |Ѫ��|0" & str_Item
        Else
            Exit Sub
        End If
        
        'ͼ������
        str_Path = App.Path & "\Gram"
        If Dir(str_Path, vbDirectory) = "" Then
            MkDir str_Path
        End If
        If InStr(str_Code, "D3U") > 0 Then
            str_imgName = "HRBC"
            str_File = Replace(str_����, "-", "_") & "_" & Val(str_�걾��) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D3U"))
        End If
        If InStr(str_Code, "D4U") > 0 Then
            str_imgName = "HPLT"
            str_File = Replace(str_����, "-", "_") & "_" & Val(str_�걾��) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D4U"))
        End If
        If InStr(str_Code, "D5U") > 0 Then
            str_imgName = "WBC DISCRI"
            str_File = Replace(str_����, "-", "_") & "_" & Val(str_�걾��) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D5U"))
        End If
        If InStr(str_Code, "D1G") > 0 Then
            str_imgName = "SDIFF"
            str_File = Replace(str_����, "-", "_") & "_" & Val(str_�걾��) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D1G"))
        End If
        If InStr(str_Code, "D2G") > 0 Then
            str_imgName = "SBASO"
            str_File = Replace(str_����, "-", "_") & "_" & Val(str_�걾��) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D2G"))
        End If
        If InStr(str_Code, "D3G") > 0 Then
            str_imgName = "SRET"
            str_File = Replace(str_����, "-", "_") & "_" & Val(str_�걾��) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D3G"))
        End If
        If InStr(str_Code, "D4G") > 0 Then
'            If bln_Rec Then
                str_imgName = "SPLT"
                str_File = Replace(str_����, "-", "_") & "_" & Val(str_�걾��) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D4G"))
'            End If
        End If
        If InStr(str_Code, "D5G") > 0 Then
            str_imgName = "SIMI"
            str_File = Replace(str_����, "-", "_") & "_" & Val(str_�걾��) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D5G"))
        End If
        If InStr(str_Code, "D6G") > 0 Then
            str_imgName = "SRET-E"
            str_File = Replace(str_����, "-", "_") & "_" & Val(str_�걾��) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D6G"))
        End If
        If InStr(str_Code, "D7G") > 0 Then
            str_imgName = "SNRBC"
            str_File = Replace(str_����, "-", "_") & "_" & Val(str_�걾��) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D7G"))
        End If
        If str_RetuImage <> "" Then strResult = strResult & vbCrLf & Mid(str_RetuImage, 2)
    Else
        'û���ݣ���ʱ������
    End If
    
End Sub

Private Function Save_image(ByVal str_img As String, ByVal str_FileName As String, ByVal str_Data As String, ByVal lng_Start As Long) As String
    Dim lng_Lens As Long, str_ImgData As String, T_ImgInfo As typHistGramInfo
    
    If lng_Start <= 0 Then Exit Function
    
    If str_img = "HRBC" Or str_img = "HPLT" Or str_img = "WBC DISCRI" Then
        lng_Lens = Val(Mid(str_Data, lng_Start + 22, 6)) - 12
        If lng_Lens <= 0 Then Exit Function
        str_ImgData = Mid(str_Data, lng_Start + 41, lng_Lens)
        T_ImgInfo.lnglower = Mid(str_Data, lng_Start + 29, 4)
        T_ImgInfo.lngupper = Mid(str_Data, lng_Start + 33, 4)
        T_ImgInfo.lngmaxx = Mid(str_Data, lng_Start + 16, 3)
        T_ImgInfo.lngmaxy = Mid(str_Data, lng_Start + 19, 3)
        T_ImgInfo.lngresver1 = 0: T_ImgInfo.lngresver2 = 0
        T_ImgInfo.lngstoppos = IIf(str_img = "PLT DISCRI", 40, 46)
        
        
        If uf_xehist(str_ImgData, lng_Lens, T_ImgInfo, "c:\temp1.bmp", str_FileName) = 1 Then
             Save_image = "^" & str_img & ";105;1;" & str_FileName
        End If
    Else
        lng_Lens = Val(Mid(str_Data, lng_Start + 22, 6)) - 1
        If lng_Lens <= 0 Then Exit Function
        str_ImgData = Mid(str_Data, lng_Start + 29, lng_Lens)
        If uf_xescat(str_ImgData, lng_Lens, "c:\Temp2.bmp", str_FileName) = 1 Then
            Save_image = "^" & str_img & ";104;1;" & str_FileName
        End If
    End If
    
End Function

Private Function GetValues(ByVal str_in As String, int_i As Integer) As String
    '����:��ʽ������ֵ
    'str_in:����ʽ�����ַ���
    'int_i:�����������ʱ,������С��λ��
    
    If str_in = "*0000" Then
        GetValues = "----"
    ElseIf str_in = "     " Then
        GetValues = " "
    Else
        GetValues = Val(str_in) / 10 ^ int_i
    End If
    
End Function