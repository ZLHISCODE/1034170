VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ABX_M60_Bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ABX Micros 60
'����:ABX
'����:�¶�
'�û�:����202ҽԺ

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
    
    Dim lngS As Long, lngE As Long, strCode As String
    Dim varline As Variant, intRow As Integer, strLine As String
    Dim str�걾�� As String, str���� As String, strItems As String, lngItem As Long
    Dim StrWbc As String, StrRbc As String, StrPLt As String, strWBCFLAG As String, strPLTFLAG As String
    Dim strBASO As String, strDiff As String, strDiffFlag As String
    Dim strimg As String
    Dim str���ڸ�ʽ As String
    Dim str�����ͺ� As String
    
    On Error GoTo errH
    
    str�����ͺ� = Trim(ReadIni("ABX_M60_Bin", "�����ͺ�", App.Path & "\ABX_M60_Bin.ini"))
    If str�����ͺ� = "" Then str�����ͺ� = "":   Call WriteIni("ABX_M60_Bin", "�����ͺ�", "", App.Path & "\ABX_M60_Bin.ini")
            
    str���ڸ�ʽ = Trim(ReadIni("ABX_M60_Bin", "���ڸ�ʽ", App.Path & "\ABX_M60_Bin.ini"))
    If str���ڸ�ʽ = "" Then str���ڸ�ʽ = "":   Call WriteIni("ABX_M60_Bin", "���ڸ�ʽ", "", App.Path & "\ABX_M60_Bin.ini")
    
    lngS = Instr_Bin(strSource, Chr(&H2))
    lngE = Instr_Bin(strSource, Chr(&H3))
    
    strReserved = strSource
    If lngS > 0 And lngE > 0 And lngE > lngS Then
        strCode = Mid_bin(strSource, lngS, lngE - lngS + 1, False)
        strReserved = Mid_bin(strSource, lngE + 1, , False)
        varline = Split(strCode, ",0D")
        For intRow = LBound(varline) To UBound(varline)
            strLine = varline(intRow)
            If InStr(",!,0,1,2,3,4,5,6,7,8,@,A,B,C,#,%,',"""",$,&,(,),*,+,-,.,/,,,", "," & Mid_bin(strLine, 1, 1) & ",") > 0 Then
                '��Ŀ
                lngItem = lngItem + 1
                strItems = strItems & "|" & lngItem & "|" & Val(Mid_bin(strLine, 3))
            ElseIf Mid_bin(strLine, 1, 1) = """" Then
                lngItem = lngItem + 1
                strItems = strItems & "|" & lngItem & "|" & Val(Mid_bin(strLine, 3))
            ElseIf Mid_bin(strLine, 1, 1) = "q" Then
                '����
                If str���ڸ�ʽ = "" Then
                    str���� = Format(CDate(Mid_bin(strLine, 9, 2) & "-" & Mid_bin(strLine, 6, 2) & "-" & Mid_bin(strLine, 3, 2)), "yyyy-MM-dd")
                Else
                    str���� = Format(CDate(Mid_bin(strLine, 3, 2) & "-" & Mid_bin(strLine, 6, 2) & "-" & Mid_bin(strLine, 9, 2)), "yyyy-MM-dd")
                End If
            ElseIf Mid_bin(strLine, 1, 1) = "s" Then
                'ABX-120 ̫ԭ��ʽ
                If str�걾�� = "" Then str�걾�� = Trim(Mid_bin(strLine, 2))
            ElseIf Mid_bin(strLine, 1, 1) = "u" Then
                '�걾��
                If Val(Trim(Replace(Mid_bin(strLine, 3), "AUTO_SID", ""))) > 0 Then
                    str�걾�� = Val(Trim(Replace(Mid_bin(strLine, 3), "AUTO_SID", "")))
                End If
                If strDiff <> "" And strDiffFlag <> "" Then strDiff = DrawBmpToFile(strDiff, "Diff", str�걾��, strDiffFlag)
                
            ElseIf Mid_bin(strLine, 1, 1) = "W" Then    'WBC
                StrWbc = Mid_bin(strLine, 2, , False)
                StrWbc = Graph_Text(StrWbc)
            ElseIf Mid_bin(strLine, 1, 1) = "X" Then    'RBC
                StrRbc = Mid_bin(strLine, 2, , False)
                StrRbc = Graph_Text(StrRbc)
            ElseIf Mid_bin(strLine, 1, 1) = "Y" Then    'PLT
                StrPLt = Mid_bin(strLine, 2, , False)
                StrPLt = Graph_Text(StrPLt)
            ElseIf Mid_bin(strLine, 1, 1) = "Z" Then    'BASO
                strBASO = Mid_bin(strLine, 2, , False)
                strBASO = Graph_Text(strBASO)
            ElseIf Mid_bin(strLine, 1, 1) = "[" Then    'DIFFPLOT
                strDiff = Mid_bin(strLine, 2, , False)
                strDiff = Graph_Diff(strDiff)
            ElseIf Mid_bin(strLine, 1, 1) = "a" Then    'DIFFPLOT Thresholds
                strDiffFlag = Trim(Mid_bin(strLine, 2))
                strDiffFlag = Replace(strDiffFlag, " ", ",")
            End If
        Next
        
        If Val(str�걾��) > 0 And strItems <> "" Then
            strResult = str���� & "|" & str�걾�� & "| |ȫѪ|0" & strItems
            
            Select Case str�����ͺ�
            Case "P80"
                If StrWbc <> "" Then strimg = strimg & "^" & "WBC;0;140,135,16.125-50,32.25-100,64.5-200,96.75-300,129-400" & StrWbc
                If StrRbc <> "" Then strimg = strimg & "^" & "RBC;0;140,135,15.45-30,51.5-100,103-200" & StrRbc
                If StrPLt <> "" Then strimg = strimg & "^" & "PLT;0;70,135,8.6-2,21.5-5,43-10,86-20,129-30" & StrPLt
                
                If strBASO <> "" Then strimg = strimg & "^" & "BAS;0;32,130,12-,25-100,38-,51-200,64-,77-300,90-,103-400,116-;" & Mid(strBASO, 2)
                If strDiff <> "" Then strimg = strimg & "^" & "LMNE;107;1;" & strDiff
            Case "P60"
                If StrWbc <> "" Then strimg = strimg & "^" & "WBC;0;140,135,16.125-50,32.25-100,64.5-200,96.75-300,129-400" & StrWbc
                If StrRbc <> "" Then strimg = strimg & "^" & "RBC;0;140,135,9-30,51.5-100,103-200" & StrRbc
                If StrPLt <> "" Then strimg = strimg & "^" & "PLT;0;70,135,8.6-2,21.5-5,43-10,86-20,129-30" & StrPLt
                
                If strBASO <> "" Then strimg = strimg & "^" & "BASO;0;70,135,43-150,86-300,129-450;" & Mid(strBASO, 2, Len(strBASO) - 2 - 4)
                If strDiff <> "" Then strimg = strimg & "^" & "V;107;1;" & strDiff
            Case Else
                If StrWbc <> "" Then strimg = strimg & "^" & "WBC;0;140,135,16.125-50,32.25-100,64.5-200,96.75-300,129-400" & StrWbc
                If StrRbc <> "" Then strimg = strimg & "^" & "RBC;0;140,135,15.45-30,51.5-100,103-200" & StrRbc
                If StrPLt <> "" Then strimg = strimg & "^" & "PLT;0;70,135,8.6-2,21.5-5,43-10,86-20,129-30" & StrPLt
                
                If strBASO <> "" Then strimg = strimg & "^" & "BASO;4;70,135,43-150,86-300,129-450;" & Mid(strBASO, 2)
                If strDiff <> "" Then strimg = strimg & "^" & "V;107;1;" & strDiff
            End Select
            If strimg <> "" Then
                strResult = strResult & vbCrLf & Mid(strimg, 2)
            End If
        End If
    End If
    
    Exit Sub
errH:
    WriteLog "M60_Bin", Err.Description, ""
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

Private Function Graph_Text(ByVal str_bin As String) As String
    '����ͼ���ı�����
    Dim lng_len As Long, lng_i As Long, str_bit As String
    Dim str_Return As String, str_in As String
    lng_len = Len(str_bin)

    If lng_len < 3 Then
        Graph_Text = ""
        Exit Function
    Else
        str_in = str_bin
    End If

    For lng_i = 1 To lng_len / 3
        
        str_bit = Replace(Left(str_in, 3), ",", "")
        str_in = Mid(str_in, 4)
        
        If IsNull(str_bit) Then
            str_Return = str_Return & "000"
        Else
            str_Return = str_Return & ";" & Format((CByte("&H" & str_bit) - 32), "000")
        End If
    Next
    Graph_Text = str_Return
End Function

Private Function Graph_Diff(ByVal str_bin As String) As String
    '����ɢ��ͼ����
    Dim lng_len As Long, str_in As String, lng_i As Long, str_bit As String
    Dim int_Stx As Integer, int_Etx As Integer, int_Pos As Integer
    Dim int_Bit As Integer, int_i As Integer
    lng_len = Len(str_bin)
    If lng_len < 3 Then
        Graph_Diff = ""
        Exit Function
    Else
        str_in = str_bin
    End If
    
    int_Bit = 1
    str_in = Mid(str_in, 4)
   Do While Len(str_in) > 0
        str_bit = Replace(Left(str_in, 3), ",", "")
        str_in = Mid(str_in, 4)
        int_Stx = CByte("&H" & str_bit)
        
        str_bit = Replace(Left(str_in, 3), ",", "")
        str_in = Mid(str_in, 4)
        int_Etx = CByte("&H" & str_bit)
            
        int_Pos = (int_Stx Mod 16) * 16 + (int_Etx Mod 16)
        If int_Pos = 0 Then
            Graph_Diff = Graph_Diff & ",00,00,00,00,00,00,00,00"
        Else
            For int_i = 1 To 8
                Graph_Diff = Graph_Diff & ",0" & (int_Pos Mod 2)
                int_Pos = Int(int_Pos / 2)
            Next
        End If
         
    Loop

End Function

Private Function DrawBmpToFile(ByVal str_bin As String, ByVal strType As String, ByVal str�걾�� As String, strFlag As String) As String
    Dim strBinData As String
    Dim lngCount As Long
    Dim strFileName As String
    
    Dim frmPic As New frmGraph
    strBinData = str_bin
    lngCount = Len(strBinData)
    If lngCount <= 0 Then Exit Function
    
    strFileName = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str�걾��) & "_" & strType & ".bmp"
    If frmPic.DrawP60(strBinData, strFileName, strFlag) Then
        DrawBmpToFile = strFileName
    End If
    Unload frmPic
    Set frmPic = Nothing
End Function