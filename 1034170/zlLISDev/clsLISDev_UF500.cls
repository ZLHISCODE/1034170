VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_UF500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: UF-500
'����: sysmex
'����: �¶�
'�û�: ����ҽԺ

Private str_Total  As String '������������
Private bln_Analyse As Boolean '�Ƿ���Կ�ʼ����
Private str_Item As String     '���������DP������
Private str_Image As String    '����������ͼ������
Private lng_count As Long       '�ظ�����ͬһ��Ϣ�Ĵ���
Private intRequest As Integer   '�ظ�����
Private Start As Variant        '��¼��ʼʱ��
Private mstrSend As String      '�ϼ�ʱ��Ӧ��ָ��.

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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�

    
    Dim lng_Etx As Long '������λ��
    Dim lng_Stx As Long '��ʼ��λ��
    Dim str_Code As String '��������
    Dim i_Count As Integer  '��Ŀ����
    Dim i As Integer 'ѭ������
    Dim blnEnd As Boolean  '�Ƿ���괫�봮
    Dim lng_Start '��ʼλ�ñ���
    Dim str�걾�� As String
    Dim str�ʿ� As String
    Dim str���� As String
    Dim var_image As Variant '��ͼ������
    Dim str_Tmp As String
    Dim str_Path As String, str_File As String, intTmp As Integer, lng_x As Long
    Dim str_RetuImage As String
'    Static intRequest As Integer
    
    strReserved = strSource
    
    Do While blnEnd = False
        If Not strReserved Like "**" & Chr(&H3) & "*" Then
            blnEnd = True
            Exit Do
        End If
        
        strCmd = Chr(&H6)
        
        lng_Etx = InStr(strReserved, Chr(&H3))
        lng_Stx = InStr(strReserved, "")
        
        str_Code = Mid(strReserved, lng_Stx + 1, lng_Etx - lng_Stx - 1)
        
        Select Case UCase(Mid(str_Code, 1, 2))
            Case "O0" '�ϼ�ʱ��ҪӦ��
                strCmd = "1|" & Chr(&H6)
'                strReserved = Replace(strReserved, "O", "Z")
                strReserved = Mid(str_Code, 24)
                mstrSend = mstrSend & "SGO" & Mid(str_Code, 10) & Chr(&H3) & "|"
                Start = Timer
                strResult = Format(Now, "YYYY-MM-DD") & "|^^0"
                Exit Do
            Case "DS" '��ʼ��
                intRequest = 0
                str_Total = str_Code
                str_Item = ""
                str_Image = ""
                bln_Analyse = False
                strReserved = Mid(strReserved, lng_Etx + 1)

                Select Case Mid(str_Total, 105, 1)
                    Case "0"
                        str_Item = "|Source|OP.CLCT"
                    Case "1"
                        str_Item = "|Source|Morning"
                    Case "2"
                        str_Item = "|Source|Timed"
                    Case "3"
                        str_Item = "|Source|AF.Meal"
                    Case "4"
                        str_Item = "|Source|Cath"
                    Case "*"
                        str_Item = "|Source|Uncertain"
                End Select
                Select Case Mid(str_Total, 106, 1)
                    Case "0"
                        str_Item = str_Item & "|UFColor|None"
                    Case "1"
                        str_Item = str_Item & "|UFColor|LyBrown"
                    Case "2"
                        str_Item = str_Item & "|UFColor|Yellow"
                    Case "3"
                        str_Item = str_Item & "|UFColor|YBrown"
                    Case "4"
                        str_Item = str_Item & "|UFColor|Orange"
                    Case "5"
                        str_Item = str_Item & "|UFColor|Red"
                    Case "6"
                        str_Item = str_Item & "|UFColor|DBrown"
                    Case "7"
                        str_Item = str_Item & "|UFColor|Green"
                    Case "8"
                        str_Item = str_Item & "|UFColor|Blue"
                    Case "9"
                        str_Item = str_Item & "|UFColor|White"
                    Case "*"
                        str_Item = str_Item & "|UFColor|Uncertain"
                End Select
                Select Case Mid(str_Total, 107, 1)
                    Case "0"
                        str_Item = str_Item & "|UFClarity|Clear"
                    Case "1"
                        str_Item = str_Item & "|UFClarity|Sle azy"
                    Case "2"
                        str_Item = str_Item & "|UFClarity|e azy"
                    Case "3"
                        str_Item = str_Item & "|UFClarity|SlCldy"
                    Case "4"
                        str_Item = str_Item & "|UFClarity|Cloudy"
                    Case "*"
                        str_Item = str_Item & "|UFClarity|Uncertain"
                End Select

                str_Item = str_Item & "|TOTAL|" & Val(Mid(str_Total, 128, 6))
                str_Item = str_Item & "|BACTOTAL|" & Val(Mid(str_Total, 134, 6))
            Case "DP"
                str_Total = str_Total & str_Code
                i_Count = Val(Mid(str_Code, 48, 2))
                For i = 0 To i_Count - 1
                    str_Item = str_Item & "|" & UCase(Mid(str_Code, 50 + 12 * i, 4))
                    str_Item = str_Item & "|" & Trim(Mid(str_Code, 54 + 12 * i, 8))
                    
                    Select Case UCase(Mid(str_Code, 50 + 12 * i, 4))
                    Case "0201"     'RBC-M
                        str_Item = str_Item & "|RBC-M|" & Val(Mid(str_Code, 54 + 12 * i, 8)) * 0.18
                    Case "0202"     'WBC-M
                        str_Item = str_Item & "|WBC-M|" & Val(Mid(str_Code, 54 + 12 * i, 8)) * 0.18
                    Case "0100"     'EC-M
                        str_Item = str_Item & "|EC-M|" & Val(Mid(str_Code, 54 + 12 * i, 8)) * 0.18
                    Case "0000"     'CAST-M
                        str_Item = str_Item & "|CAST-M|" & Val(Mid(str_Code, 54 + 12 * i, 8)) * 2.9
                    Case "0401"     'BACT-M
                        str_Item = str_Item & "|BACT-M|" & Val(Mid(str_Code, 54 + 12 * i, 8)) / 100
                    End Select
                Next
                strReserved = Mid(strReserved, lng_Etx + 1)
            Case "DC"
                str_Total = str_Total & str_Code
                strReserved = Mid(strReserved, lng_Etx + 1)
            Case "DQ"
                str_Total = str_Total & str_Code
                i_Count = Val(Mid(str_Code, 48, 2))
                For i = 0 To i_Count - 1
                    str_Item = str_Item & "|" & UCase(Mid(str_Code, 50 + 12 * i, 4))
                    str_Item = str_Item & "|" & Trim(Mid(str_Code, 54 + 12 * i, 8))
                    
                    Select Case UCase(Mid(str_Code, 50 + 12 * i, 4))
                    Case "00D9" 'P.CAST  PATH.CAST 0-1
                        If Val(Trim(Mid(str_Code, 54 + 12 * i, 8))) >= 0 And Val(Trim(Mid(str_Code, 54 + 12 * i, 8))) <= 1 Then
                            str_Item = str_Item & "|PATH.CAST|-"
                        Else
                            str_Item = str_Item & "|PATH.CAST|+"
                        End If
                    Case "0107" 'SRC SRCF           0-3
                        If Val(Trim(Mid(str_Code, 54 + 12 * i, 8))) >= 0 And Val(Trim(Mid(str_Code, 54 + 12 * i, 8))) <= 3 Then
                            str_Item = str_Item & "|SRCF|-"
                        Else
                            str_Item = str_Item & "|SRCF|+"
                        End If
                    Case "0501" 'SPERM   SPERMF     ?
                    
                    Case "0300" 'X,TAL   XTALF      0-15
                        If Val(Trim(Mid(str_Code, 54 + 12 * i, 8))) >= 0 And Val(Trim(Mid(str_Code, 54 + 12 * i, 8))) <= 15 Then
                            str_Item = str_Item & "|XTALF|-"
                        Else
                            str_Item = str_Item & "|XTALF|+"
                        End If
                    Case "0402" 'YLC YLCF           0-10
                        If Val(Trim(Mid(str_Code, 54 + 12 * i, 8))) >= 0 And Val(Trim(Mid(str_Code, 54 + 12 * i, 8))) <= 10 Then
                            str_Item = str_Item & "|YLCF|-"
                        Else
                            str_Item = str_Item & "|YLCF|+"
                        End If
                    Case "00DA" 'MUCUS   MUCUSF     ?
                    Case "0502" 'Cond.   Cond.F     5-38
                        If Val(Trim(Mid(str_Code, 54 + 12 * i, 8))) >= 5 And Val(Trim(Mid(str_Code, 54 + 12 * i, 8))) <= 38 Then
                            str_Item = str_Item & "|Cond.F|-"
                        Else
                            str_Item = str_Item & "|Cond.F|+"
                        End If
                    End Select
                Next
                strReserved = Mid(strReserved, lng_Etx + 1)
            Case "DD"
                str_Total = str_Total & str_Code
                i_Count = Val(Mid(str_Code, 48, 2))
                For i = 0 To i_Count - 1
                    If UCase(Mid(str_Code, 50 + 12 * i, 4)) = "0C00" Then
                        Select Case Trim(Mid(str_Code, 54 + 12 * i, 8))
                            Case "00000000"
                                str_Item = str_Item & "|0C00|RBC Negative"
                            Case "00000001"
                                str_Item = str_Item & "|0C00|Isomorphic?"
                            Case "00000002"
                                str_Item = str_Item & "|0C00|Dysmorphic?"
                            Case "00000003"
                                str_Item = str_Item & "|0C00|Mixed?"
                        End Select
                    ElseIf UCase(Mid(str_Code, 50 + 12 * i, 4)) = "0C01" Then
                        Select Case Trim(Mid(str_Code, 54 + 12 * i, 8))
                            Case "00000000"
                                str_Item = str_Item & "|0C01|Not flagged"
                            Case "00000001"
                                str_Item = str_Item & "|0C01|RANK1"
                            Case "00000002"
                                str_Item = str_Item & "|0C01|RANK2"
                            Case "00000003"
                                str_Item = str_Item & "|0C01|RANK3"
                            Case "00000004"
                                str_Item = str_Item & "|0C01|RANK4"
                            Case "00000005"
                                str_Item = str_Item & "|0C01|RANK5"
                        End Select
                    ElseIf UCase(Mid(str_Code, 50 + 12 * i, 4)) = "0C02" Then
                        Select Case Trim(Mid(str_Code, 54 + 12 * i, 8))
                            Case "00000000"
                                str_Item = str_Item & "|0C02|Not flagged"
                            Case "00000001"
                                str_Item = str_Item & "|0C02|UTI?"
                        End Select
                    Else
                        str_Item = str_Item & "|" & UCase(Mid(str_Code, 50 + 12 * i, 4))
                        str_Item = str_Item & "|" & Trim(Mid(str_Code, 54 + 12 * i, 8))
                    End If
                Next
                strReserved = Mid(strReserved, lng_Etx + 1)
            Case "DV"
                str_Total = str_Total & str_Code
                i_Count = Val(Mid(str_Code, 48, 2))
                For i = 0 To i_Count - 1
                    str_Item = str_Item & "|" & UCase(Mid(str_Code, 50 + 12 * i, 4))
                    str_Item = str_Item & "|" & Trim(Mid(str_Code, 54 + 12 * i, 8))
                Next
                strReserved = Mid(strReserved, lng_Etx + 1)
            Case "DF"
                str_Total = str_Total & str_Code
                i_Count = Val(Mid(str_Code, 48, 2))
                For i = 0 To i_Count - 1
                    str_Item = str_Item & "|" & UCase(Mid(str_Code, 50 + 12 * i, 4))
                    str_Item = str_Item & "|" & Trim(Mid(str_Code, 54 + 12 * i, 8))
                Next
                strReserved = Mid(strReserved, lng_Etx + 1)
            Case "DI"
                'ͼ������
                If Mid(str_Code, 5, 2) = "01" Then
                    str_Image = str_Image & "|" & Mid(str_Code, 91, 16) & "^" & Mid(str_Code, 115)
                Else
                    str_Image = str_Image & Mid(str_Code, 48)
                End If
                str_Total = str_Total & str_Code
                strReserved = Mid(strReserved, lng_Etx + 1)
            Case "D1"
                If UCase(Mid(str_Code, 1, 3)) = "D1C" Then
                    bln_Analyse = True
                    str_Total = str_Total & str_Code
                    strReserved = Mid(strReserved, lng_Etx + 1)
                End If
            Case "DE" '������
                bln_Analyse = True
                str_Total = str_Total & str_Code
                strReserved = Mid(strReserved, lng_Etx + 1)
                If InStr(strReserved, "") <= 0 Then strReserved = ""
            Case Else
                str_Total = str_Total & str_Code
                strReserved = Mid(strReserved, lng_Etx + 1)
        End Select
        
        If bln_Analyse Then
            If InStr(str_Total, "DS") > 0 Then
                str�ʿ� = "0"
            Else
                If InStr(str_Total, "D1C") > 0 Then
                    str�ʿ� = "1"
                    'D1C��Ϊ�赥������,����������,�ݲ�����
                    'call D1C
                    GoTo Lable_0
                End If
            End If
            lng_Start = InStr(str_Total, "DS")
            
            str�걾�� = Mid(str_Total, lng_Start + 70, 15)
            If str�ʿ� <> "1" Then
                str�ʿ� = IIf(Mid(str_Total, lng_Start + 10, 1) = "C", "1", "0")
            End If
                        
            str���� = Mid(str_Total, lng_Start + 48, 4)
            str���� = str���� & "-" & Mid(str_Total, lng_Start + 52, 2)
            str���� = str���� & "-" & Mid(str_Total, lng_Start + 54, 2)
            If IsDate(str����) Then
                str���� = Format(str����, "yyyy-MM-dd")
            Else
                str���� = Format(Now, "yyyy-MM-dd")
            End If
            
                            
            If Val(str�걾��) > 0 And str_Item <> "" Then
                str_RetuImage = ""
                strResult = str���� & "|" & Val(str�걾��) & "| |��Һ|0" & str_Item
                If str_Image <> "" Then
                    
                    str_Image = Mid(str_Image, 2)
                    var_image = Split(str_Image, "|")
                    
                    str_Path = App.Path & "\Gram"
                    If Dir(str_Path, vbDirectory) = "" Then
                        MkDir str_Path
                    End If
                    
                    For i = LBound(var_image) To UBound(var_image)
                        str_File = ""
                        Do While str_File = ""
                            intTmp = intTmp + 1
                            str_File = Replace(str����, "-", "") & "_" & Format(Val(str�걾��), "000") & "_" & Format(intTmp, "00") & ".gif"
                            If Dir(str_Path & "\" & str_File) <> "" Then str_File = ""
                        Loop
                        
                        str_Tmp = var_image(i)
                        If InStr(str_Tmp, "WBC") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                            If uf_uf1000Hist(str_Tmp, Len(str_Tmp), "c:\tempfile1.bmp", str_Path & "\" & str_File) = 1 Then
                                str_RetuImage = str_RetuImage & "^WBC Fsc;104;4;" & str_Path & "\" & str_File
                            End If
                        ElseIf InStr(str_Tmp, "RBC") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                            If uf_uf1000Hist(str_Tmp, Len(str_Tmp), "c:\tempfile1.bmp", str_Path & "\" & str_File) = 1 Then
                                str_RetuImage = str_RetuImage & "^RBC Fsc;104;4;" & str_Path & "\" & str_File
                            End If
                        ElseIf InStr(str_Tmp, "SPERM") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
'                            If uf_uf1000Hist(str_Tmp, Len(str_Tmp), "c:\tempfile1.bmp", str_Path & "\" & str_File) = 1 Then
'                                str_RetuImage = str_RetuImage & "^SPERM Fsc;100;1;" & str_Path & "\" & str_File
'                            End If
                        ElseIf InStr(str_Tmp, "S_FLHxS_FSC") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                            If uf_uf1001scat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 128, 128) = 1 Then
                                str_RetuImage = str_RetuImage & "^FLHxS-FSC;104;4;" & str_Path & "\" & str_File
                            End If
                        ElseIf InStr(str_Tmp, "S_FLLxS_FSC") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                            If uf_uf1001scat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 128, 128) = 1 Then
                                str_RetuImage = str_RetuImage & "^FLLxS-FSC;104;4;" & str_Path & "\" & str_File
                            End If
                        ElseIf InStr(str_Tmp, "S_FLLWxS_FLLW2") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                            If uf_uf1000scat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 128, 128) = 1 Then
                                str_RetuImage = str_RetuImage & "^FLLWx-FLLW2;104;4;" & str_Path & "\" & str_File
                            End If
                        ElseIf InStr(str_Tmp, "S_SSCxS_FSC") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                            If uf_uf1001scat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 128, 128) = 1 Then
                                str_RetuImage = str_RetuImage & "^SSCxS-FSC;104;1;" & str_Path & "\" & str_File
                            End If
                        ElseIf InStr(str_Tmp, "S_FSCWxS_FLLW") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                            If uf_uf1000scat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 128, 128) = 1 Then
                                str_RetuImage = str_RetuImage & "^FSCWxS-FLLW;104;4;" & str_Path & "\" & str_File
                            End If
                        ElseIf InStr(str_Tmp, "SEB_FLHxB_FSC") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                            If uf_uf1000scat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 128, 128) = 1 Then
                                str_RetuImage = str_RetuImage & "^FLHxB-FSC;104;1;" & str_Path & "\" & str_File
                            End If
                        ElseIf InStr(str_Tmp, "SEB_FSCWxB_FSC") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
'                            If uf_uf1000scat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 128, 128) = 1 Then
'                                str_RetuImage = str_RetuImage & "^FSCWxB-FSC;100;1;" & str_Path & "\" & str_File
'                            End If
                        ElseIf InStr(str_Tmp, "SEB_FLHxB_FSC2") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
'                            If uf_uf1000scat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 128, 128) = 1 Then
'                                str_RetuImage = str_RetuImage & "^FLHxB-FSC2;100;1;" & str_Path & "\" & str_File
'                            End If
                            
                        End If
                    Next
                    If str_RetuImage <> "" Then strResult = strResult & vbCrLf & Mid(str_RetuImage, 2)
                    
                End If
                
            End If
    
            '--- ��ձ���

            str_Total = ""
            str_Item = ""
            str_Image = ""
            bln_Analyse = False
            blnEnd = True
        End If

        
    Loop '���괫�봮
'    If strReserved <> "" And strCmd = "" Then strCmd = Chr(&H6)
    
    Exit Sub
Lable_0:
    str_Total = ""
    str_Item = ""
    bln_Analyse = False
    blnEnd = True
End Sub

Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    If InStr(mstrSend, "|") > 0 Then
        SendSample = Split(mstrSend, "|")(0)
        mstrSend = Mid(mstrSend, InStr(mstrSend, "|") + 1)
    End If

End Function

