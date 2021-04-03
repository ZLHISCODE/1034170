VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Roche_P800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'����:�ٴ���������ϵ P800
'����:Roche ����
'����:�¶�
'�û�:��ú������ҽԺ

Private mstrCode As String    '��������
Private mlngStep_In As Long   '�����ַ�����������

Private mlngStep_Out As Long  '��˫����Ϣ����
Private mstrRack As String    'ɨ����ʱ�����̺�,�ܺŵ���Ҫ���ص���Ϣ.

Private mstrWaitSend As String '�����ͱ걾
Private mstrQCNO As String      '��QC��ţ�ҪתΪ�̶��ı걾��

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
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    Dim lng_S As Long, lng_E As Long, i_Count As Integer
    Dim str_in As String, str_Line As String, str_���� As String, var_line As Variant
    Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String, str_��� As String
    Dim str_���� As String, str_�ʿ� As String
    '-----------------------------------------------------------------------------
    Dim varQC As Variant, strIniFile As String, strQC As String, intloop As Integer
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream, strLine As String
    
    str_in = strSource
        
    If InStr(str_in, Chr(&H4)) > 0 Then                 'EOT '����ͨѶ����
        
        mlngStep_In = 0
    End If
    
    If InStr(str_in, Chr(&H3)) > 0 Then                 '�յ�ETX(03)
        lng_S = InStr(str_in, Chr(&H3))
        mlngStep_Out = mlngStep_Out + 1
        strCmd = Chr(&H6)
    End If
    
    If InStr(str_in, Chr(&H5)) > 0 Then                 '�յ�ENQ(05) Ӧ�� ETX(03)
        strCmd = Chr(&H6)
    End If
    
    If InStr(str_in, Chr(&H17)) > 0 Then                  '�յ�ETB(17) Ӧ��
        strCmd = Chr(&H6)
    End If
    
    
    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    'ȥ��
    Do While mstrCode Like "*" & Chr(&H17) & "*" & Chr(&HD) & Chr(&HA) & Chr(&H2) & "*"
        lng_S = InStr(mstrCode, Chr(&H17))
        lng_E = InStr(lng_S, mstrCode, Chr(&H2)) + 1
        mstrCode = Mid(mstrCode, 1, lng_S - 1) & Mid(mstrCode, lng_E + 1)
    Loop
    
    lng_S = InStr(mstrCode, Chr(&H5))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    Else
        lng_E = InStr(mstrCode, Chr(&H4))
    End If
    If lng_S > 0 And lng_E > 0 Then
        str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        If InStr(str_Line, "|TSREQ^REAL|") > 0 Then         '˫��ͨѶ- ��ʼ
            mstrRack = ""
            mlngStep_Out = 1
            var_line = Split(str_Line, Chr(&HD))
            
            For i_Count = LBound(var_line) To UBound(var_line)
                str_Line = var_line(i_Count)
                If str_Line Like "Q|*" Then
                    str_���� = Trim(Split(str_Line, "|")(2))
                    
                    mstrRack = Split(str_����, "^")(3) & "^" & Split(str_����, "^")(4) & "^" & Split(str_����, "^")(5) & "^^" & Split(str_����, "^")(7) & "^" & Split(str_����, "^")(8)
                    str_���� = Split(str_����, "^")(2)
                    
                    If Val(str_����) <> 0 Then
                        If Len(CStr(Val(str_����))) >= 12 Then
                            strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & Trim(str_����)
'                        Else
                            '�걾��
'                            strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|" & Trim(str_����)
                        End If
                    End If
                End If
            Next
            
            If strResult <> "" Then
                strResult = Mid(strResult, 3)
                strCmd = "1|"       '��������ͨѶ��ʼ���� 05
            End If

        ElseIf InStr(str_Line, "|RSUPL^") > 0 Then     '������
        '-----------------------------------------------
            strIniFile = App.Path & "\Roche_P800.ini"
            strQC = ""
            If objFileSystem.FileExists(strIniFile) Then
                Set objStream = objFileSystem.OpenTextFile(strIniFile)
                Do While Not objStream.AtEndOfStream
                    strLine = objStream.ReadLine
                    If InStr(strLine, "=") > 0 Then
                        strQC = strQC & "|" & strLine
                    End If
                Loop
'            Else
'               Call objFileSystem.CreateTextFile(strIniFile)
'               Set objStream = objFileSystem.OpenTextFile(strIniFile, ForAppending)
'               objStream.WriteLine "[ˮƽ1]"
'               objStream.WriteLine "5200 = PPU"
'               objStream.WriteLine "[ˮƽ2]"
'               objStream.WriteLine "6100 = PNU"
'
'               strQC = "5200 = PPU|6100 = PNU"
            End If
            If strQC <> "" Then varQC = Split(Mid(strQC, 2), "|")
            '-----------------------------------------------------------------------------
            var_line = Split(str_Line, Chr(&HD))
            
            For i_Count = LBound(var_line) To UBound(var_line)
                str_Line = var_line(i_Count)
                If Left(str_Line, 1) = Chr(&HA) Then str_Line = Mid(str_Line, 2)
                If str_Line Like "O|*" Then
                    str_���� = Trim(Split(str_Line, "|")(7))
                    If Len(str_����) <> 14 Then str_���� = Split(str_Line, "|")(22)
                    str_���� = Left(str_����, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2)
                    
                    str_�걾�� = Split(str_Line, "|")(2)
                    
                    str_���� = Trim(Split(str_Line, "|")(5))
                    str_�ʿ� = IIf(Trim(Split(str_Line, "|")(11)) = "Q", "1", "0")
                    
                    If str_�ʿ� = "1" Then
                        If strQC <> "" Then
                            For intloop = LBound(varQC) To UBound(varQC)
                                If InStr("," & UCase(Trim(Split(varQC(intloop), "=")(1))) & ",", "," & UCase(Trim(str_�걾��)) & ",") > 0 Then
                                    str_�걾�� = " " & Val(Split(varQC(intloop), "=")(0))
                                    Exit For
                                End If
                            Next
                            End If
                        End If
                    If InStr(str_�걾��, "^") > 0 Then
                        ' �����ܺ�ȫ�ıߵĸ�ʽ
                        str_�걾�� = Split(str_�걾��, "^")(0)
                    End If
                    If Trim(str_�걾��) <> "" And Len(Trim(str_�걾��)) < 12 Then
                            '�걾��
                        str_�걾�� = Trim(str_�걾��) & IIf(str_���� = "S", "^1", "")
                    ElseIf Trim(str_�걾��) <> "" And Val(Trim(str_�걾��)) > 0 Then
                            '����
                        str_�걾�� = "^" & IIf(str_���� = "S", "^1", "^") & Trim(str_�걾��)
                    Else
                        str_�걾�� = ""
                    End If
                    
                    str_��Ŀ = ""
                ElseIf str_Line Like "R|*" Then
                    str_��� = Trim(Split(str_Line, "|")(3))
                    If str_��� <> "" Then
                        If InStr(str_���, "^") > 0 Then str_��� = Split(str_���, "^")(1) '--- 090407 �����ߵĽ����Ϊ -1^0.715 �ĸ�ʽ��
                        str_��Ŀ = str_��Ŀ & "|" & Replace(Replace(Split(str_Line, "|")(2), "^^^", ""), "/", "") & "|" & str_���
                    End If
                End If
            Next
            If IsDate(CDate(str_����)) And str_�걾�� <> "" And str_��Ŀ <> "" Then
                strResult = str_���� & "|" & str_�걾�� & "| |Ѫ��|" & str_�ʿ� & str_��Ŀ
                WriteLog "�յ��������صĽ�� ", strResult, ""
                '���� ������Ŀ
                var_line = Split(Mid(str_��Ŀ, 2), "|")
                
                For intloop = LBound(var_line) To UBound(var_line)
                    If Trim(var_line(intloop)) = "320" Then 'Cysc
                        If Val(var_line(intloop + 1)) <> 0 Then
                            strResult = strResult & "|4513|" & (78 / Val(var_line(intloop + 1)) + 4)
                            WriteLog "������ ", "|4513|" & (78 / Val(var_line(intloop + 1)) + 4), ""
                            Exit For
                        End If
                    End If
                Next
                If Val(Getarr(str_��Ŀ, "678")) <> 0 And Val(Getarr(str_��Ŀ, "413")) <> 0 Then
                    strResult = strResult & "|GLO|" & Val(Getarr(str_��Ŀ, "678")) - Val(Getarr(str_��Ŀ, "413"))
                End If
                
                If Val(Getarr(str_��Ŀ, "18")) <> 0 And Val(Getarr(str_��Ŀ, "294")) <> 0 Then
                    strResult = strResult & "|IBIL|" & Val(Getarr(str_��Ŀ, "18")) - Val(Getarr(str_��Ŀ, "294"))
                End If
                
                If Val(Getarr(str_��Ŀ, "320")) <> 0 And Val(Getarr(str_��Ŀ, "311")) <> 0 Then
                    strResult = strResult & "|X001|" & Val(Getarr(str_��Ŀ, "320")) / Val(Getarr(str_��Ŀ, "311"))
                End If
            End If

        End If
    End If
    
End Sub


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
    Dim str_WaitSendLine As String
    
    blnSuccess = True
    If iSendStep = 0 Then
        If strSampleInfo <> "" Then WriteLog "˫��ͨѶ(P800 Begin)��", strSampleInfo, mstrWaitSend
        var_�걾 = Split(strSampleInfo, "||")
        
        SendSample = ""
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
            
            str_������Ŀ = ""
            
            For i_Loop = 10 To UBound(Split(str_��¼, "|"))
                If Val(Split(str_��¼, "|")(i_Loop)) > 0 Then
                    If InStr(str_������Ŀ, "^^^" & Replace(Replace(Split(str_��¼, "|")(i_Loop), "/1/not", ""), "1not", "") & "^\") <= 0 Then
                        '����Ѫ�ǣ�ֻ��һ��
                        str_������Ŀ = str_������Ŀ & "^^^" & Replace(Replace(Split(str_��¼, "|")(i_Loop), "/1/not", ""), "1not", "") & "^\"
                    End If
                End If
            Next
            
            If str_������Ŀ <> "" And str_���� <> "" Then
                str_������Ŀ = Mid(str_������Ŀ, 1, Len(str_������Ŀ) - 1)
                
                str_SendItem = "H|\^&|||H7600^1||||||TSDWN^REPLY" & Chr(&HD) & "P|1" & Chr(&HD)
                str_SendItem = str_SendItem & "O|1|" & str_���� & Space(24 - Len(str_����))
                str_SendItem = str_SendItem & "|" & mstrRack & "|" & str_������Ŀ
                str_SendItem = str_SendItem & "|" & IIf(i_���� = 1, "S", "R")               'S-����,R-����
                str_SendItem = str_SendItem & "||" & Format(Now, "yyyyMMddHHmmss")
                str_SendItem = str_SendItem & "||||N||||1||||||||||O" & Chr(&HD)
                str_SendItem = str_SendItem & "C|1|L|^^^^|G" & Chr(&HD)
                str_SendItem = str_SendItem & "L|1|N" & Chr(&HD) & Chr(&H3)
                          
                str_EndString = ""
                lngRow = 0
                Do While str_SendItem <> ""
                    
                    If Len(str_SendItem) <= 230 Then
                        lngRow = lngRow + 1
                        str_EndString = str_EndString & "," & Chr(&H2) & lngRow & str_SendItem & GetCheckCode(lngRow & str_SendItem) & vbCrLf
                        str_SendItem = ""
                    Else
                        lngRow = lngRow + 1
                        str_Tmp = Mid(str_SendItem, 1, 230) & Chr(&H17)
                        str_EndString = str_EndString & "," & Chr(&H2) & lngRow & str_Tmp & GetCheckCode(lngRow & str_Tmp) & vbCrLf
                        str_SendItem = Mid(str_SendItem, 231)
                    End If
                Loop
    
                If SendSample = "" Then
                    SendSample = SendSample & Chr(&H5)                          '���� ��Ϣ������־(EOT)
                    mstrWaitSend = mstrWaitSend & str_EndString & "," & Chr(&H4)
                Else
                    mstrWaitSend = mstrWaitSend & Chr(&H5) & str_EndString & "," & Chr(&H4)
                End If
                iSendStep = 1
            End If
        Next
        If strSampleInfo <> "" Then WriteLog "˫��ͨѶ(P800 End)��", SendSample, mstrWaitSend
    Else
        str_Tmp = strResponse
        If InStr(str_Tmp, Chr(&H6)) > 0 Then                 '�յ�ACK(06)
            If mstrWaitSend <> "" Then
                If Left(mstrWaitSend, 1) = "," Then
                    mstrWaitSend = Mid(mstrWaitSend, 2)
                End If
                If InStr(mstrWaitSend, ",") > 0 Then
                    str_WaitSendLine = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
                    mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))
                Else
                    str_WaitSendLine = mstrWaitSend
                    mstrWaitSend = ""
                End If
                
                SendSample = str_WaitSendLine
                
            End If
        End If
        If mstrWaitSend = "" Then iSendStep = 0
    End If
End Function


Private Function GetCheckCode(ByVal str_in As String) As String

'��ȡУ����
    Dim lngAsc As Long, i As Integer
    
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function

'Public Function GetTool() As String
'    '�������߽ӿ�,������Ҫȡ�ʿ�Ʒ��Ϣ
'    GetTool = "GetQCInfo"
'End Function
'
'Public Function DoTool(ByVal strInMessage As String) As String
'    '��ʾ���ô���
'    frmSetup_Roche_P800.mstrInfo = strInMessage
'    frmSetup_Roche_P800.Show vbModal
'End Function
Private Function Getarr(strLine As String, strFind As String) As String
    'ȡ�����е�ֵ
    Dim intRow As Integer
    Dim astrLine() As String
    astrLine = Split(strLine, "|")
    For intRow = 0 To UBound(astrLine)
        If astrLine(intRow) = strFind Then
            Getarr = astrLine(intRow + 1)
        End If
    Next
End Function
