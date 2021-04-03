VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Elec2010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'����:���߷��� 2010 ˫��
'����:Roche ����
'����:�¶�
'�û�:��ú������ҽԺ

Private mstrCode     As String    '��������

Private mlngStep_In  As Long   '�����ַ�����������

Private mlngStep_Out As Long  '��˫����Ϣ����

Private mstrRack     As String    'ɨ����ʱ�����̺�,�ܺŵ���Ҫ���ص���Ϣ.

Private mstrWaitSend As String '�����ͱ걾

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

Public Sub Analyse(ByVal strSource As String, _
                   strResult As String, _
                   strReserved As String, _
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

    Dim str_WaitSendLine As String, str_���� As String, str_�ʿ� As String

    Dim str_Last�걾�� As String

    str_in = strSource
        
    If InStr(str_in, Chr(&H4)) > 0 Then                 'EOT '����ͨѶ����
        mlngStep_In = 0
    End If
    
    If InStr(str_in, Chr(&H3)) > 0 Then                 '�յ�ETX(03)
        lng_S = InStr(str_in, Chr(&H3))
        mlngStep_Out = mlngStep_Out + 1
        strCmd = Chr(&H6)
    End If
    
    '    If InStr(str_in, Chr(&H6)) > 0 Then                 '�յ�ACK(06)
    '        lng_S = InStr(str_in, Chr(&H6))
    '        str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)
    '        strReserved = str_in
    '        str_in = ""
    '    End If
    
    If InStr(str_in, Chr(&H5)) > 0 And strCmd = "" Then               '�յ�ENQ(05) Ӧ�� ETX(03)
        strCmd = Chr(&H6)
    End If
    
    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
    lng_S = InStr(mstrCode, Chr(&H5))
    lng_E = InStr(mstrCode, Chr(&H4))

    If lng_S > 0 And lng_E > 0 Then
        
        str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        WriteLog "��ʼ", "�յ���������", str_Line
        
        If str_Line Like "*" & Chr(&H2) & "?Q|*" Then          '˫��ͨѶ- ��ʼ
            mstrRack = ""
            mlngStep_Out = 1
            var_line = Split(str_Line, Chr(&HA))
            
            For i_Count = LBound(var_line) To UBound(var_line)
                str_Line = var_line(i_Count)

                If str_Line Like Chr(&H2) & "?Q|*SAMPLE^NORMAL||ALL||||||||*" Then   'ɨ����
                    str_���� = Split(str_Line, "|")(2)
                    
                    mstrRack = Split(str_����, "^")(2) & "^" & Split(str_����, "^")(3) & "^" & Split(str_����, "^")(4)
                    str_���� = Split(str_����, "^")(1)
                    
                    If Val(str_����) <> 0 Then
                        strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & Trim(str_����)
                        
                    End If
                End If

            Next
            
            If strResult <> "" Then
                strResult = Mid(strResult, 3)
                strCmd = "1|" & strCmd     '��������ͨѶ��ʼ���� 05
                WriteLog "����", "������", strResult
            End If

        ElseIf str_Line Like "*" & Chr(&H2) & "?R|*" Then     '������
            var_line = Split(str_Line, Chr(&HA))
            
            For i_Count = LBound(var_line) To UBound(var_line)
                str_Line = var_line(i_Count)

                If str_Line Like Chr(&H2) & "?O|*" Then
                    str_���� = Split(str_Line, "|")(6)
                    str_���� = Left(str_����, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2)
                    
                    str_�걾�� = Split(str_Line, "|")(2)
                    str_���� = Trim(Split(str_Line, "|")(5))
                    
                    str_�ʿ� = "0"
                    
                    If Trim(Split(str_Line, "|")(11)) = "Q" Then str_�ʿ� = "1"

                    '---- ��ú���ʿر걾
                    If Trim(Split(str_Line, "|")(11)) = "X\Q" Then str_�ʿ� = "1"
                    If Trim(str_�걾��) Like "PC*1" Then
                        str_�ʿ� = "1"
                        str_�걾�� = "3100"
                    ElseIf Trim(str_�걾��) Like "PC*2" Then
                        str_�ʿ� = "1"
                        str_�걾�� = "3200"
                    End If
                    
                    If Val(Trim(str_�걾��)) > 0 Then
                        If Len(Trim(str_�걾��)) < 12 Then
                            '�걾��
                            str_�걾�� = Trim(str_�걾��) & IIf(str_���� = "S", "^1", "")
                        Else
                            '����
                            str_�걾�� = "^" & IIf(str_���� = "S", "^1", "^") & Trim(str_�걾��)
                        End If

                    Else
                        str_�걾�� = ""
                    End If
                    
                    str_��Ŀ = ""
                ElseIf str_Line Like Chr(&H2) & "?R|*" Then
                    str_��� = Trim(Split(str_Line, "|")(3))

                    If str_��� <> "" Then
                        If InStr(str_���, "^") > 0 Then
                            str_��Ŀ = "|" & Split(Split(str_Line, "|")(2), "^")(3) & "|" & Split(str_���, "^")(1)
                        Else
                            str_��Ŀ = "|" & Split(Split(str_Line, "|")(2), "^")(3) & "|" & str_���
                        End If

                        If IsDate(CDate(str_����)) And str_�걾�� <> "" And str_��Ŀ <> "" Then
                            If str_Last�걾�� = str_�걾�� Then
                                strResult = strResult & str_��Ŀ
                            Else
                                strResult = strResult & "||" & str_���� & "|" & str_�걾�� & "| |Ѫ��|" & str_�ʿ� & str_��Ŀ
                                str_Last�걾�� = str_�걾��
                            End If
                        End If
                    End If
                End If

            Next

            If strResult <> "" Then
                If Val(Getarr(strResult, "391")) <> 0 And Val(Getarr(strResult, "321")) <> 0 Then
                    
                    strResult = strResult & "|X003|" & Val(Getarr(strResult, "391")) / Val(Getarr(strResult, "321"))
                End If
                strResult = Mid(strResult, 3)
            End If
            WriteLog "����", "������", strResult
        End If
    End If
    
End Sub

Public Function SendSample(strSampleInfo As String, _
                           iSendStep As Integer, _
                           blnSuccess As Boolean, _
                           Optional ByVal strResponse As String = "", _
                           Optional ByVal blnUndo As Boolean = False, _
                           Optional ByVal iType As Integer = 0) As String
    
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

    Dim str_SendItem As String, str_Line As String
    
    WriteLog "��ʼ", strSampleInfo, ""
    
    var_�걾 = Split(strSampleInfo, "||")
    SendSample = ""
    
    blnSuccess = True

    If iSendStep = 0 Then

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
                    str_������Ŀ = str_������Ŀ & "^^^" & Split(str_��¼, "|")(i_Loop) & "^0\"
                End If

            Next
            
            If str_������Ŀ <> "" And str_���� <> "" Then
                str_������Ŀ = Mid(str_������Ŀ, 1, Len(str_������Ŀ) - 1)
                
                str_Line = "1H|\^&|||ASTM-Host" & Chr(&HD) & Chr(&H3)
                str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","
                
                str_Line = "2P|1||" & str_���� & Chr(&HD) & Chr(&H3)
                str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","
                
                str_Line = "3O|1|" & str_���� & "|" & mstrRack & "|" & str_������Ŀ & "|" & IIf(i_���� = 1, "S", "R")
                str_Line = str_Line & "||||||N||||||||||||||O" & Chr(&HD) & Chr(&H3)
                str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","
                
                str_Line = "4L|1|" & Chr(&HD) & Chr(&H3)
                str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","
                
                If SendSample = "" Then
                    SendSample = SendSample & Chr(&H5)                              '���� ��Ϣ������־(EOT)
                    iSendStep = 1                                                   '��ʼͨ��
                    mstrWaitSend = mstrWaitSend & "," & str_SendItem & Chr(&H4)
                Else
                    mstrWaitSend = mstrWaitSend & "," & Chr(&H5) & "," & str_SendItem & Chr(&H4)
                End If
            End If

        Next

    Else

        If InStr(strResponse, Chr(&H6)) > 0 Then
            If mstrWaitSend <> "" Then
                
                If Left(mstrWaitSend, 1) = "," Then
                    mstrWaitSend = Mid(mstrWaitSend, 2)
                End If

                If InStr(mstrWaitSend, ",") > 0 Then
                    SendSample = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
                    mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))
                Else
                    SendSample = mstrWaitSend
                    mstrWaitSend = ""
                End If
                
                WriteLog "�յ�06,���͵�������ָ��", SendSample, ""
                
            End If
        End If

        If mstrWaitSend = "" Then iSendStep = 0

        Exit Function

    End If
    
    WriteLog "����", "���͵���������Ϣ", SendSample
    
End Function

Private Function GetCheckCode(ByVal str_in As String) As String

    '��ȡУ����
    Dim lngAsc As Long, i As Integer
    
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
    
End Function
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
