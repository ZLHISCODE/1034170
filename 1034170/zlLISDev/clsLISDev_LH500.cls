VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_LH500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrCode As String
Const conPicPath = "\draw.bmp"
'����:LH-500/LH750
'2007-9-28���ֵڶ�����ҽԺ: tqh����
'�޸�ͼ���㷨
'�޸��ˣ�������
'2012-11-20
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

Public Sub Analyse(ByVal strSource As String, strResult As String, Reserved As String, _
                   strCmd As String)
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    '����ͼ��ķ�ʽ��
    '            1.ͼ�����ݸ���ָ�����ݺ�ʹ�ûس����з����ָ���
    '            2.�ж��ͼ������ʱʹ��"^"���ָ�
    '            3.����ͼ�����ݸ�ʽ: ͼ�񻭷� 0=ֱ��ͼ  1=ɢ��ͼ
    '              a) ֱ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ);X1;X2;X3;X4;X5...
    '              b) ɢ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ):
    '                 ��:00000100001000010000100010;00000100001000010000100010;
    '                 ˵��:1.ɢ��ͼ�Ե���ʽ����ÿһ��ʹ�÷ֺ����ָ�.
    '                      2.�ж��ٸ��ֺž��ж�����
    '                      3.ÿһ���ж��ٸ�����ÿһ�еĳ�����ȷ��
    '                      3.��ͼ�ķ����Ǵ����ϱ����»�������65*65��ͼ���Ǵ�65�п�ʼ��(���ϱ߿�ʼ��)
    Dim str_Line As String
    Dim lngS As Long        '��ʼ����
    Dim lngE As String      '��������
    Dim strDate As String   '����
    Dim strNo As String     '�걾��
    Dim strItems As String   '������Ŀ
    Dim StrWbc As String    'ֱ��ͼ
    Dim StrRbc As String    '�����ֱ��ͼ
    Dim StrPLt As String
    Dim strDiff As String   'ɢ��ͼ
    Dim strImgs As String   'ͼ���ַ���
    Dim strSample As String
    Dim strDIFFLine As String
    Dim varDIFF As Variant
    Dim varItem As Variant
    Dim strLine As String
    Dim i As Integer
    Dim j As Integer
    Dim intloop As Integer
    Dim strNowData As Integer
    Dim str_���� As String
    Dim str_getData As String
    Dim str_ͼ������ As String
    str_ͼ������ = ";RBCH;PLTH;PLTF;WBCT;WBCF;DIFF;"
    mstrCode = mstrCode & strSource

    mstrCode = Replace(mstrCode, "", "|")    ' ȥ���ַ����ڵ��ַ�
    lngE = 0
    lngS = InStr(mstrCode, "")
    If lngS > 0 Then
        lngE = InStr(lngS, mstrCode, "")
    Else
        lngE = InStr(mstrCode, "")
    End If
    Do While lngS > 0 And lngE > 0
        strSample = Mid(mstrCode, lngS, lngE - lngS + 1)
        mstrCode = Mid(mstrCode, lngE + 1)
        varItem = Split(strSample, vbCrLf)
        For i = LBound(varItem) To UBound(varItem)
            strLine = varItem(i)
            If Len(strLine) > 0 Then
                Select Case Split(strLine, " ")(0)
                Case "DATE"
                    strDate = Format(Split(strLine, " ")(1), "YYYY-MM-DD")
                Case "TIME"
                    strDate = strDate & " " & Split(strLine, " ")(1)
                    '                    strDate = Format(strDate, "yyyy-mm-dd hh;mm:ss")
                Case "ID1"
                    strNo = Split(strLine, " ")(1)
                Case "CASSPOS"
                    If Val(strNo) = 0 Then
                        str_���� = Split(strLine, " ")(1)
                        For j = 1 To Len(str_����)
                            If IsNumeric(Left(str_����, 1)) Then
                                Exit For
                            Else
                                str_���� = Mid(str_����, 2)
                            End If
                        Next
                    End If
                Case "WBC", "RBC", "HGB", "MCV", "MCH", "RDW", "PLT", "MPV", "PDW", "PCT", "MO#", "NE#", "EO#", "BA#", "LY%", "MO%", "NE%", "EO%", "BA%"
                    strItems = strItems & "|" & Split(strLine, " ")(0) & "|" & Split(strLine, " ")(1)
                Case "HCT"
                    strItems = strItems & "|" & Split(strLine, " ")(0) & "|" & Mid(Split(strLine, " ")(1), 1, 5)
                Case "RBCH", "PLTH", "WBCT", "WBCF"  '����RBCHͼ
                    strImgs = strImgs & "^" & getStrImg(Trim(Split(strLine, " ")(0)), Trim(Split(strLine, " ")(1)))
                Case "PLTF"
                    strImgs = strImgs & "|" & getStrImg(Trim(Split(strLine, " ")(0)), Trim(Split(strLine, " ")(1)))
                Case "DIFF"
                    If Len(strLine) > 4900 Then    '����ɢ��ͼ
                        strImgs = strImgs & "^" & getStrImg("DIFF", Trim(Split(strLine, " ")(1)))
                    End If
                End Select
            End If
        Next
        If strDate <> "" And strNo <> "" And strImgs <> "" And strItems <> "" Then
            If str_���� <> "" Then
                strResult = strResult & "||" & strDate & "|" & Val(strNo) & "^^" & str_���� & "| |Ѫ��|0" & strItems & IIf(strImgs <> "", vbCrLf & Mid(strImgs, 2), "")
            Else
                strResult = strResult & "||" & strDate & "|" & Val(strNo) & "| |Ѫ��|0" & strItems & IIf(strImgs <> "", vbCrLf & Mid(strImgs, 2), "")
            End If
        End If
        lngS = InStr(mstrCode, "")
        If lngS > 0 Then
            lngE = InStr(lngS, mstrCode, "")
        Else
            lngE = InStr(mstrCode, "")
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

Public Function getStrImg(ByVal strType As String, ByVal strData As String) As String
    Dim str_R As String
    Dim varDIFF As Variant
    Dim strItem As String
    Dim strLine As String
    Dim j As Integer
    Select Case strType
    Case "DIFF"    '����DIFFͼ������
        varDIFF = Split(strData, "|")
        For j = 1 To UBound(varDIFF) - 1
            str_R = str_R & Mid(varDIFF(j), 3, Len(varDIFF(j)) - 2 - 4)
        Next
        '        Debug.Print str_R
    Case Else    '����ͼ������
        If InStr(strData, "|") > 0 Then
            varDIFF = Split(strData, "|")
            strLine = Mid(varDIFF(0), 1, Len(varDIFF(0)) - 4)
            For j = 1 To UBound(varDIFF) - 1
                strLine = strLine & Mid(varDIFF(j), 3, Len(varDIFF(j)) - 2 - 4)
            Next
            strLine = strLine & Mid(varDIFF(UBound(varDIFF)), 3)
        Else
            strLine = strData
        End If

        Do While Len(strLine) > 0
            strItem = Left(strLine, 2)
            strLine = Mid(strLine, 3)
            If CByte("&H" & strItem) >= 0 Then
                If InStr(strType, "PLT") > 0 Then
                    str_R = str_R & "," & CByte("&H" & strItem)
                Else
                    str_R = str_R & ";" & CByte("&H" & strItem)
                End If
            End If
        Loop
    End Select
    If str_R <> "" Then
        Select Case strType
        Case "RBCH"
            getStrImg = "RBC;0;32,350,52-50,104-100,156-,208-200,260-,302-300" & str_R
        Case "WBCT"
            getStrImg = "WBCT;0;32,70,15-30,30-60" & str_R
        Case "WBCF"
            getStrImg = "WBCF;0;100,250,20-50,50-100,100-,150-200" & str_R
        Case "PLTH"
            getStrImg = "PLT;4;100,200,4-2,50-10,100-20,150-30;" & Mid(str_R, 2)
        Case "PLTF"
            getStrImg = Mid(str_R, 2)
        Case "DIFF"
            getStrImg = "DIFF;100;1;" & drawDIFF(str_R)
        End Select
    End If
End Function

Public Function drawDIFF(ByVal strImgData As String) As String
    Dim x As Integer
    Dim Y As Integer
    Dim strLine As String
    Dim strTmp As String
    frmGraph.Picture1.Scale (0, 0)-(256, 256)
    frmGraph.Picture1.BackColor = vbWhite
    strLine = strImgData
    '    Debug.Print Len(strLine)
    For Y = 1 To 64
        For x = 1 To 64
            If Mid(strLine, 1, 2) <> "00" And Mid(strLine, 1, 2) <> "" Then
                Call DrawPoint(Mid(strLine, 1, 2), x, Y)
            End If
            strLine = Mid(strLine, 3)
        Next
    Next
    If Dir(App.Path & conPicPath) <> "" Then
        Kill App.Path & conPicPath
    End If
    SavePicture frmGraph.Picture1.Image, App.Path & conPicPath
    drawDIFF = App.Path & conPicPath
End Function
Public Function DrawPoint(ByVal str_in As String, ByVal x As Integer, ByVal Y As Integer)
    Dim strColor As String
    Dim i As Integer
    Dim j As Integer
    frmGraph.Picture1.DrawWidth = 1
    Select Case str_in
    Case "01", "02", "03", "04", "05", "06", "07", "08", "09"
        frmGraph.Picture1.PSet (x * 4, Y * 4), vbWhite
    Case "0B"    '��ɫ
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H40C0&
    Case "0C"
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &H80FF&
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H80FF&
    Case "0D"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H80FF&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &H40C0&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H80FF&
    Case "0E"
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H80FF&
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H80FF&
    Case "0F"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H80FF&
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &H80FF&

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &H80FF&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H40C0&
    Case "10"
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H80FF&
    Case "11"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H80FF&
    Case "15"    'ˮ��ɫ
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &HFF00FF
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &HFF00FF
    Case "16", "17", "18", "19"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &HFF00FF
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &HFF00FF
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HC000C0
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HFF00FF

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &HC000C0
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &HFF00FF
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &HFF00FF

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &H800080
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HFF00FF
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &HC000C0

    Case "1A"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H400040
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &H400040
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HC000C0
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HFF00FF

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &HC000C0
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H400040
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &HFF00FF

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &H800080
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &H400040
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &HC000C0

    Case "1B"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H400040
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &H400040
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HC000C0
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HFF00FF

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &HC000C0
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H400040
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &HFF00FF

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &H800080
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &H400040
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &HC000C0

    Case "1C", "1D"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H400040
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &H800080
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HC000C0
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &H400040

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &HC000C0
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H400040
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &H400040

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &H400040
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &H400040
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &H800080
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &HC000C0

    Case "1F"    '��ɫ
        frmGraph.Picture1.PSet (x * 3, Y * 4), &HFF00&
    Case "20"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &HFF00&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HC000&
    Case "21"
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &HFF00&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &H8000&
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HFF00&
    Case "22"
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &HC000&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &H8000&
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HC000&
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H8000&
    Case "23", "24"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H8000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &HFF00&
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &H8000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HC000&

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &HC000&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H8000&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF00&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &HC000&

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HFF00&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HC000&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &H8000&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HFF00&
    Case "25"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &HFF00&
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &H8000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &H4000&
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HC000&

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &HFF00&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H4000&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HC000&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &H8000&

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HFF00&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HC000&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HFF00&
        frmGraph.Picture1.PSet (x * 3 - 3, Y * 4 - 3), &H4000&

    Case "29"    '����
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HFF0000
    Case "2A"
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HC00000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HC00000
    Case "2C"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H0
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HC00000
    Case "2B"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HFF8080
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HFF0000
    Case "2D"
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &HC00000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HFF0000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HFF0000

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 3), &HFF0000

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HFFFF00
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HFF8080
    Case "2E"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &HFF0000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &HFF0000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HC00000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HFF0000

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &HC00000

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HC00000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &HFF0000
    Case "2F"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &HFF8080
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &HFF0000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HC00000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &H800000

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &H800000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &HC00000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &HFF8080

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HFF8080
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HFF8080
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &H800000

    Case Else
        '            Debug.Print str_in
    End Select

End Function

