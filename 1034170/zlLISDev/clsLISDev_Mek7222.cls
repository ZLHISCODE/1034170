VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Mek7222_Bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Ѫ���� Mek 7222
'����:�ձ����
'����:�¶�
'�û�:�Ͼ�����ǿ
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
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    Dim varline As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    Dim strImage As String  'ͼ������
    Dim varImage As Variant 'ͼ������
    Dim strɢ������ As String, strɢ1 As String, strɢ2 As String, strɢ3 As String
    Dim lng_S As Long, lng_E As Long
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, ",02")
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, ",03")
    Else
        lng_E = InStr(mstrCode, ",03")
    End If

    If lng_E > 0 And lng_S > 0 Then
        lng_E = lng_E + 4608 * 3 'ֱ��ͼ����
        lng_E = lng_E + 40000 * 3 'ɢ��ͼ����
        lng_E = lng_E + 112 * 3 'ɢ��ͼ��������
        If Len(Mid(mstrCode, lng_S)) < (lng_E - lng_S) + 3 Then lng_E = 0
    End If
    
    Do While lng_S > 0 And lng_E > 0
    
        lng_E = InStr(lng_S, mstrCode, ",03")

        'ȡһ�������ļ�¼
        strLine = Mid(mstrCode, lng_S, lng_E - lng_S + 3)
        strImage = Mid(mstrCode, lng_E + 3, 4608 * 3)
        
        strɢ������ = Mid(mstrCode, lng_E + 4608 * 3 + 3, 40112 * 3)
  
        lng_E = lng_E + 4608 * 3 'ֱ��ͼ����
        lng_E = lng_E + 40000 * 3 'ɢ��ͼ����
        lng_E = lng_E + 112 * 3 'ɢ��ͼ��������
        
        mstrCode = Mid(mstrCode, lng_E + 3)
        
        If Len(strImage) = 4608 * 3 And Len(strɢ������) = 40112 * 3 Then
            WriteLog "�յ���������", strLine & vbNewLine & strImage & vbNewLine & "ɢ�㣺" & strɢ������, mstrCode
            strLine = Mid_bin(strLine, 1)
            strImage = Mid_bin(strImage, 1)
            
            varline = Split(strLine, Chr(&HD))
            If strImage <> "" Then
                varImage = Split(strImage, Chr(&HD))
                If UBound(varImage) < 768 Then
                    strLine = ""
                    lng_S = 0: lng_E = 0
                End If
            End If
        Else
            lng_S = 0: lng_E = 0
        End If

        If lng_S > 0 And lng_E > 0 Then
            str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = "": strTmp = ""
            '����
            
            '�������
            For iRow = 0 To UBound(varline)
                Select Case iRow
                Case 0 '����
                    If varline(iRow) Like "*##-*-##*" Then
                        str���� = Mid(varline(iRow), InStr(varline(iRow), "-") - 2, 2)
                        strTmp = Mid(varline(iRow), InStr(varline(iRow), "-") + 1)
                        strTmp = Mid(strTmp, InStr(strTmp, "-") + 1, 2) & "-" & TrunMonth(Mid(strTmp, 1, InStr(strTmp, "-") - 1))
                        str���� = Format(CDate(strTmp & "-" & str����), "yyyy-mm-dd")
                    ElseIf varline(iRow) Like "*##/???/##*" Then
                        str���� = Mid(varline(iRow), InStr(varline(iRow), "/") - 2, 2)
                        strTmp = Mid(varline(iRow), InStr(varline(iRow), "/") + 1)
                        strTmp = Mid(strTmp, InStr(strTmp, "/") + 1, 2) & "-" & TrunMonth(Mid(strTmp, 1, InStr(strTmp, "/") - 1))
                        str���� = Format(CDate(strTmp & "-" & str����), "yyyy-mm-dd")
                    Else
                        str���� = Format(Now, "yyyy-MM-dd")
                    End If
                Case 1 '�걾��
                    str�걾�� = Val(varline(iRow))
                    str�����˱걾�ʿ� = " |ѪҺ|0"
                Case 2 'WBC  ��ϸ������
                    strItem = "|WBC|" & Format_6318(varline(iRow))
                Case 3 'LYM% �ܰ�ϸ���ٷֱ�
                    strItem = strItem & "|LY%|" & Format_6318(varline(iRow))
                Case 4 'MON% ��ֵϸ���ٷֱ�
                    strItem = strItem & "|MO%|" & Format_6318(varline(iRow))
                Case 5 'GRA% ����ϸ���ٷֱ�
                    strItem = strItem & "|NE%|" & Format_6318(varline(iRow))
                Case 6 'LYM# �ܰ�ϸ������ֵ
                    strItem = strItem & "|EO%|" & Format_6318(varline(iRow))
                Case 7 'MON# ��ֵϸ������ֵ
                    strItem = strItem & "|BA%|" & Format_6318(varline(iRow))
                Case 8 'GRA# ����ϸ������ֵ
                    strItem = strItem & "|LY|" & Format_6318(varline(iRow))
                Case 9 'EOS# ��������ϸ������ֵ
                    strItem = strItem & "|MO|" & Format_6318(varline(iRow))
                Case 10 'RBC  ��ϸ������
                    strItem = strItem & "|NE|" & Format_6318(varline(iRow))
                Case 11 'HGB Ѫ�쵰��
                    strItem = strItem & "|EO|" & Format_6318(varline(iRow))
                Case 12 'HCT  ��ϸ��ѹ��
                    strItem = strItem & "|BA|" & Format_6318(varline(iRow))
                Case 13 'MCV  ƽ����ϸ�����
                    strItem = strItem & "|RBC|" & Format_6318(varline(iRow))
                Case 14 'MCH  ƽ��Ѫ�쵰�׺���
                    strItem = strItem & "|HGB|" & Format_6318(varline(iRow))
                Case 15 'MCHC ƽ��Ѫ�쵰��Ũ��
                    strItem = strItem & "|HCT|" & Format_6318(varline(iRow))
                Case 16 'RDW  ��ϸ��ƽ������
                    strItem = strItem & "|MCV|" & Format_6318(varline(iRow))
                Case 17 'PLT ѪС������
                    strItem = strItem & "|MCH|" & Format_6318(varline(iRow))
                Case 18 'PCT  ѪС��ѹ��
                    strItem = strItem & "|MCHC|" & Format_6318(varline(iRow))
                Case 19 'MPV  ƽ��ѪС�����
                    strItem = strItem & "|RDW|" & Format_6318(varline(iRow))
                Case 20 'PDW  ѪС��ƽ������
                    strItem = strItem & "|PLT|" & Format_6318(varline(iRow))
                Case 21 'PDW  ѪС��ƽ������
                    strItem = strItem & "|PCT|" & Format_6318(varline(iRow))
                Case 22 'PDW  ѪС��ƽ������
                    strItem = strItem & "|MPV|" & Format_6318(varline(iRow))
                Case 23 'PDW  ѪС��ƽ������
                    strItem = strItem & "|PDW|" & Format_6318(varline(iRow))
                End Select
            Next
            
            '����ͼ��
            If strImage <> "" Then
                For iRow = 0 To UBound(varImage)
                    If iRow >= 0 And iRow <= 256 Then
                        If iRow = 0 Then strImage = "^WBC;0;0,271,15-50,51-100,87-,123-200,159-,195-300,235-,271-400"
                        strImage = strImage & ";" & Val(varImage(iRow))
                    End If
                    If iRow > 256 And iRow <= 512 Then
                        If iRow = 257 Then strImage = strImage & "^RBC;0;0,256,18-,56-100,94-,132-200,170-,208-300,246-"
                        strImage = strImage & ";" & Val(varImage(iRow))
                    End If
                    If iRow > 512 And iRow <= 768 Then
                        If iRow = 513 Then strImage = strImage & "^PLT;0;0,220,9-2,42-,75-20,108-,141-40,174-,207-60"
                        strImage = strImage & ";" & Val(varImage(iRow))
                    End If
                Next
            End If
            
            If strItem <> "" And Val(str�걾��) <> 0 Then
                strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
                
                '����ɢ��ͼ
                If strɢ������ <> "" Then
                    strɢ1 = SaveBmp(str����, str�걾��, strɢ������)
                    If strɢ1 <> "" Then
                        strImage = strImage & strɢ1
                    End If
                End If
                
                If strImage <> "" Then
                    strResult = strResult & vbCrLf & Mid(strImage, 2)
                End If
            End If
    
            lng_S = InStr(mstrCode, ",02")
            If lng_S > 0 Then
                lng_E = InStr(lng_S, mstrCode, ",03")
            Else
                lng_E = InStr(mstrCode, ",03")
            End If
        
            If lng_E > 0 And lng_S > 0 Then
                lng_E = lng_E + 4608 * 3 'ֱ��ͼ����
                lng_E = lng_E + 40000 * 3 'ɢ��ͼ����
                lng_E = lng_E + 112 * 3 'ɢ��ͼ��������
                If Len(Mid(mstrCode, lng_S)) < (lng_E - lng_S) + 3 Then lng_E = 0
            End If
        End If
    Loop


    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If

End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:������Ҫ���������͵ı걾��Ϣ
    '           �걾֮����||�ָ�
    '           Ԫ��֮����|�ָ�
    '           ��0��Ԫ�أ�����ʱ��
    '           ��1��Ԫ�أ��������
    '           ��2��Ԫ�أ�������
    '           ��3��Ԫ�أ��걾
    '           ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '           ��5��9Ԫ�أ�ϵͳ����
    '           �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    '   iSendStep  : ����
    '   blnSuccess : �Ƿ���ʧ��
    '   strResponse: ������Ӧ����Ϣ
    '   blnUndo    : �Ƿ�ȡ��.
    '
    '   ����ֵ�������������ִ���ָ�
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function


Private Function TrunMonth(ByVal strMonth As String) As String
    'ת��Ӣ�ĵ��·�Ϊ����
    Select Case Mid(UCase(strMonth), 1, 3)
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select

End Function

Private Function Format_6318(ByVal strVal As String) As String
    '��ʽת��
    '��>,<��ԭ�����أ�����תΪ���ַ���
    If InStr(strVal, ">") > 0 Or InStr(strVal, "<") > 0 Then
        Format_6318 = Trim(strVal)
    Else
        Format_6318 = Val(Trim(strVal))
    End If
End Function



Private Function SaveBmp(ByVal str���� As String, ByVal str�걾�� As String, ByVal strData As String) As String
    '����3��ɢ��ͼ��BMP�ļ�
    Dim frmPic1 As New frmGraph, strFile1 As String
    Dim frmPic2 As New frmGraph, strFile2 As String
    Dim frmPic3 As New frmGraph, strFile3 As String
    
    Dim strɢ1 As String, strɢ2 As String, strɢ3 As String, strTmp As String
    Dim lngX1 As Long, lngX2 As Long, lngY1 As Long, lngY2 As Long
    Dim lngloop As Long
    Dim strIn As String
    
    strIn = strData
    If Len(strIn) <> 40112 * 3 Then Exit Function
    
    frmPic1.Picture1.Scale (0, 0)-(256, 256)
    frmPic1.Picture1.BackColor = vbWhite
    
    frmPic2.Picture1.Scale (0, 0)-(256, 256)
    frmPic2.Picture1.BackColor = vbWhite
    
    frmPic3.Picture1.Scale (0, 0)-(256, 256)
    frmPic3.Picture1.BackColor = vbWhite
    
    For lngloop = 1 To 10000
        strTmp = Mid(strIn, 1, 12)
        strɢ1 = CByte("&H" & Mid(strTmp, 2, 2)) 'Size��ʾѪϸ���Ĵ�С
        strɢ2 = CByte("&H" & Mid(strTmp, 5, 2)) 'Complexity��ʾѪϸ���ĸ��Ӷ�
        strɢ3 = CByte("&H" & Mid(strTmp, 8, 2)) 'Granularity��ʾѪϸ���Ŀ����ĳ̶�
        strIn = Mid(strIn, 13)
        '������Size��������Complexity��ɢ��ͼ�����³�ΪS-Cɢ��ͼ����
        frmPic1.Picture1.PSet (Val(strɢ2), 256 - Val(strɢ1)), &H80000011  's-cͼ
        '������Size��������Granularity�������԰�ϸ��/�����԰�ϸ��������ɢ��ͼ�����³�ΪNE-EO������ɢ��ͼ��
        If Val(strɢ3) < 35 And Val(strɢ2) < 48 Then
            frmPic3.Picture1.PSet (Val(strɢ3), 256 - Val(strɢ1)), &H80000011  'NE-EOͼ
        End If
        If Val(strɢ3) >= 36 Or Val(strɢ2) >= 50 Then
        '������Size��������Granularity�ĵ���ϸ��/�ȼ��԰�ϸ��������ɢ��ͼ�����³�ΪMO-BA������ɢ��ͼ����
            frmPic2.Picture1.PSet (Val(strɢ3), 256 - Val(strɢ1)), &H80000011  'MO-BAͼ
        End If
    Next
    strɢ1 = ""
    strFile1 = App.Path & "\" & Replace(str����, "-", "") & "_" & str�걾�� & "_S1.JPG"
    strFile2 = App.Path & "\" & Replace(str����, "-", "") & "_" & str�걾�� & "_S2.JPG"
    strFile3 = App.Path & "\" & Replace(str����, "-", "") & "_" & str�걾�� & "_S3.JPG"
    
    'pic1������
    lngX1 = 255: lngY1 = 44
    lngX2 = 128: lngY2 = 44
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 128: lngY1 = 44
    lngX2 = 65: lngY2 = 90
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 65: lngY1 = 90
    lngX2 = 50: lngY2 = 90
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 20: lngY1 = 255
    lngX2 = 20: lngY2 = 70
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 21: lngY1 = 70
    lngX2 = 30: lngY2 = 58
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 30: lngY1 = 58
    lngX2 = 35: lngY2 = 56
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 36: lngY1 = 56
    lngX2 = 50: lngY2 = 71
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 50: lngY1 = 71
    lngX2 = 50: lngY2 = 81
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 52: lngY1 = 81
    lngX2 = 52: lngY2 = 102
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 54: lngY1 = 102
    lngX2 = 54: lngY2 = 152
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 54: lngY1 = 152
    lngX2 = 90: lngY2 = 255
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)
    
    'pic2������
    lngX1 = 255: lngY1 = 44
    lngX2 = 0: lngY2 = 44
    frmPic2.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 0: lngY1 = 132
    lngX2 = 82: lngY2 = 132
    frmPic2.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 82: lngY1 = 132
    lngX2 = 177: lngY2 = 44
    frmPic2.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 82: lngY1 = 132
    lngX2 = 250: lngY2 = 255
    frmPic2.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)
    
    'pic3������
    lngX1 = 255: lngY1 = 60
    lngX2 = 0: lngY2 = 60
    frmPic3.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)
    
    lngX1 = 255: lngY1 = 150
    lngX2 = 0: lngY2 = 150
    frmPic3.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)
    
    lngX1 = 50: lngY1 = 255
    lngX2 = 50: lngY2 = 60
    frmPic3.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)
'
'    SavePicture frmPic1.Picture1.Image,  APP.Path "\TMP1.BMP"
'    SavePicture frmPic2.Picture1.Image, strFile2
'    SavePicture frmPic3.Picture1.Image, strFile3
    
    SavePic frmPic1.Picture1.Image, strFile1, "JPG"
    SavePic frmPic2.Picture1.Image, strFile2, "JPG"
    SavePic frmPic3.Picture1.Image, strFile3, "JPG"
    
    SaveBmp = "^Wbc1;123;1;" & strFile1 & "^Wbc2;123;1;" & strFile2 & "^Wbc3;123;1;" & strFile3
    Exit Function

ErrHand:
    SaveBmp = ""
End Function