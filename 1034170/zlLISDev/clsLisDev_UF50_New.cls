VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_UF50_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: UF-50
'����: sysmex
'����: �¶�
'�û�: ��ҽ��Ժ

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
        strReserved = Replace(strReserved, "", "")
        
        lng_Etx = InStr(strReserved, Chr(&H3))
        lng_Stx = InStr(strReserved, "")
        
        str_Code = Mid(strReserved, lng_Stx + 1, lng_Etx - lng_Stx - 1)
        
'        If InStr(str_Code, "") Then lng_Stx = InStr(str_Code, "")
'        str_Code = Mid(str_Code, lng_Stx + 1, lng_Etx - lng_Stx - 1)
        
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
                '���̵ĸ�ʽ��Ϊ������ʽһ��.
                'str_Total = Left(str_Total, 27)  Right(str_Total, Len(str_Total) - 27)
                Select Case Mid(str_Total, 50, 1)
                    Case "0"
                        'str_Item = "|RBCInfo|RBC Negative"
                        str_Item = "|RBCI|��"
                    Case "1"
                        'str_Item = "|RBCInfo|Normocytic"
                        str_Item = "|RBCI|Normocytic?"
                    Case "2"
                        'str_Item = "|RBCInfo|Microcytic"
                        str_Item = "|RBCI|Microcytic?"
                    Case "3"
                        'str_Item = "|RBCInfo|Non-classified"
                        str_Item = "|RBCI|Non Classified?"
                End Select
                If IsNumeric(Mid(str_Total, 51, 5)) Then
                    str_Item = str_Item & "|COND|" & Val(Mid(str_Total, 51, 5))
                Else
                    str_Item = str_Item & "|COND|" & Mid(str_Total, 51, 5)
                End If
                '��ҽ��Ժ 12.02Ҫ�󣬴���10000,���ϱ�10000���ϵ�����C
                str_Item = str_Item & "|TOTAL|" & Val(Mid(str_Total, 91, 6)) 'Val(Mid(str_Total, 91, 6)))
            Case "DP"
                str_Total = str_Total & str_Code
                i_Count = Val(Mid(str_Code, 9, 2))
                For i = 0 To i_Count - 1
                    str_Item = str_Item & "|" & UCase(Mid(str_Code, 11 + 12 * i, 4))
                    If IsNumeric(Mid(str_Code, 15 + 12 * i, 8)) Then
                        str_Item = str_Item & "|" & Val(Mid(str_Code, 15 + 12 * i, 8))
                    Else
                        str_Item = str_Item & "|" & Trim(Mid(str_Code, 15 + 12 * i, 8))
                    End If
                    
                    Select Case UCase(Mid(str_Code, 11 + 12 * i, 4))
                    Case "0201"     'RBC-M
                        str_Item = str_Item & "|WBC-M|" & Val(Mid(str_Code, 15 + 12 * i, 8)) * 0.18
                    Case "0202"     'WBC-M
                        str_Item = str_Item & "|RBC-M|" & Val(Mid(str_Code, 15 + 12 * i, 8)) * 0.18
                    Case "0100"     'EC-M
                        str_Item = str_Item & "|EC-M|" & Val(Mid(str_Code, 15 + 12 * i, 8)) * 0.18
                    Case "0000"     'CAST-M
                        str_Item = str_Item & "|CAST-M|" & Val(Mid(str_Code, 15 + 12 * i, 8)) * 2.92
                    Case "0401"     'BACT-M
                        str_Item = str_Item & "|BACT-M|" & Val(Mid(str_Code, 15 + 12 * i, 8)) * 0.18
                    
                    End Select
                    
                Next
                strReserved = Mid(strReserved, lng_Etx + 1)
            Case "DC"
                str_Total = str_Total & str_Code
                strReserved = Mid(strReserved, lng_Etx + 1)
            Case "DI"
                'ͼ������
                
                If Mid(str_Code, 5, 2) = "01" Then
                    str_Image = str_Image & "|" & Mid(str_Code, 79, 10) & "^" & Mid(str_Code, 95)
                Else
                    str_Image = str_Image & Mid(str_Code, 30)
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
            
            str�걾�� = Mid(str_Total, lng_Start + 27, 15)
            If str�ʿ� <> "1" Then
                str�ʿ� = IIf(Mid(str_Total, lng_Start + 10, 1) = "C", "1", "0")
            End If
                        
            str���� = Mid(str_Total, lng_Start + 9, 2)
            str���� = str���� & "-" & Mid(str_Total, lng_Start + 11, 2)
            str���� = str���� & "-" & Mid(str_Total, lng_Start + 13, 2)
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
                            If uf_ufHist(str_Tmp, Len(str_Tmp), "c:\tempfile1.bmp", str_Path & "\" & str_File) = 1 Then
                                str_RetuImage = str_RetuImage & "^WBC Fsc;105;1;" & str_Path & "\" & str_File
                            End If
                        ElseIf InStr(str_Tmp, "RBC") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                            If uf_ufHist(str_Tmp, Len(str_Tmp), "c:\tempfile1.bmp", str_Path & "\" & str_File) = 1 Then
                                str_RetuImage = str_RetuImage & "^RBC Fsc;105;1;" & str_Path & "\" & str_File
                            End If
                        ElseIf InStr(str_Tmp, "Fscw-Flw") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                            If uf_ufscat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 256, 128) = 1 Then
                                str_RetuImage = str_RetuImage & "^Fscw-Flw;104;1;" & str_Path & "\" & str_File
                            End If
                        ElseIf InStr(str_Tmp, "Fl2-Fsc") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                            If uf_ufscat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 256, 128) = 1 Then
                                str_RetuImage = str_RetuImage & "^Fl2-Fsc;104;1;" & str_Path & "\" & str_File
                            End If
                        ElseIf InStr(str_Tmp, "Fl-Fsc") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                            If uf_ufscat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 128, 128) = 1 Then
                                str_RetuImage = str_RetuImage & "^Fl-Fsc;104;1;" & str_Path & "\" & str_File
                            End If
                        ElseIf InStr(str_Tmp, "Fscw-Fl") > 0 Then
                            str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                            If uf_ufscat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 128, 128) = 1 Then
                                str_RetuImage = str_RetuImage & "^Fscw-Fl;104;1;" & str_Path & "\" & str_File
                            End If
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

