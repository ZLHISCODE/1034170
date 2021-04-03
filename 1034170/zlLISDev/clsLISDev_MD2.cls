VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MD2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'����:�ٴ���������ϵ MD2 (DXC800/LX20)
'����:��������������� Beckman coulter
'����:�¶�
'�û�:������˾����

Private mstrCode As String    '��������
Private mlngStep_In As Long   '�����ַ�����������
Private mstrACK As String     '�Ƿ�Ӧ��ACK,����Ӧ��ACK(06),��-Ӧ�� ETX(03)

Private mlngStep_Out As Long  '��˫����Ϣ����
Private mstr������ As String

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
    Dim str_In As String, str_Line As String, str_���� As String
    Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String, str_��� As String
    
    str_In = strSource
    
    If InStr(str_In, Chr(&H4) & Chr(&H1)) > 0 Then       'EOT SOH '����ͨѶ���� Ӧ��
        lng_S = InStr(str_In, Chr(&H4) & Chr(&H1))
        str_In = Mid(str_In, 1, lng_S - 1) & Mid(str_In, lng_S + 2)
        strCmd = ACKChar
        
        If mlngStep_In <> 0 Then WriteLog "����" & mlngStep_In, "�յ���ʼ(EOT SOH)����.", ""
        mlngStep_In = 1
    End If
        
    If InStr(str_In, Chr(&H4)) > 0 Then                 'EOT '����ͨѶ����
        lng_S = InStr(str_In, Chr(&H4))
        str_In = Mid(str_In, 1, lng_S - 1) & Mid(str_In, lng_S + 1)
        If mlngStep_In = 0 Then WriteLog "����" & mlngStep_In, "�յ�����(EOT)����.", ""
        mlngStep_In = 0
        Call ACKChar
        strCmd = Chr(&H4)
    End If
    
    If InStr(str_In, Chr(&H5)) > 0 Then                 '�յ�ENQ(05) Ӧ�� ETX(03)
        lng_S = InStr(str_In, Chr(&H5))
        str_In = Mid(str_In, 1, lng_S - 1) & Mid(str_In, lng_S + 1)
        strCmd = Chr(&H3)
    End If
    
    If InStr(str_In, Chr(&HA)) > 0 Then                 '�յ�����(0A) Ӧ��
        lng_S = InStr(str_In, Chr(&HA))
        str_In = Mid(str_In, 1, lng_S - 1) & Mid(str_In, lng_S + 1)
        strCmd = ACKChar
    End If
    
    If Len(str_In) > 0 Then mstrCode = mstrCode & str_In
    
    lng_S = InStr(mstrCode, "[")
    lng_E = InStr(mstrCode, "]")
    If lng_S > 0 And lng_E > 0 And InStr(mstrCode, Chr(&HD)) > 0 Then
        str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        lng_E = InStr(mstrCode, Chr(&HD))
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        If InStr(str_Line, ",801,06,") > 0 Then         '˫��ͨѶ- ��ʼ
            mstr������ = ""
            mlngStep_Out = 1
            For i_Count = 0 To 3
                str_���� = Split(str_Line, ",")(3 + i_Count)
                str_���� = Trim(Replace(str_����, "]", ""))
                If str_���� <> "" Then
                    strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & str_����
                End If
            Next
            If strResult <> "" Then
                mstr������ = Split(str_Line, ",")(0)
                strResult = Mid(strResult, 3)
                Call ACKChar
                strCmd = "1|" & ACKChar & Chr(&H4) & Chr(&H1)     '��������ͨѶ��ʼ���� EOT SOH
            End If
        ElseIf InStr(str_Line, ",801,02,") > 0 Then     '��������ȷ����Ϣ

        ElseIf InStr(str_Line, ",802,01,") > 0 Then     '������ʼ���ؽ��

        ElseIf InStr(str_Line, ",802,05,") > 0 Then     '�������ؽ�����

        ElseIf InStr(str_Line, ",802,03,") > 0 Then     '������
            str_���� = Split(str_Line, ",")(3)
            str_���� = Right(str_����, 4) & "-" & Mid(str_����, 3, 2) & "-" & Left(str_����, 2)
            str_�걾�� = Trim(Split(str_Line, ",")(9))
            str_��Ŀ = Trim(Split(str_Line, ",")(10))
            str_��� = Trim(Split(str_Line, ",")(15))
            strResult = str_���� & "|^^" & str_�걾�� & "| |Ѫ��|0|" & str_��Ŀ & "|" & str_���

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
    Dim i_���� As Integer, str_���� As String, str_������Ŀ As String, str_�걾���� As String, str������ As String
    Dim str_SendItem As String, str_WaitSendLine As String, str_In As String
    
    blnSuccess = True
    If iSendStep = 0 Then
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
                str_������Ŀ = str_������Ŀ & "," & Split(str_��¼, "|")(i_Loop) & " ,0"
            Next
            
            
            If str_������Ŀ <> "" And str_���� <> "" Then
                str_������Ŀ = Format(i_Loop - 10, "000") & str_������Ŀ
                If mstr������ = "" Then
                    If Dir(App.Path & "\DM2.ini") <> "" Then
                        str������ = ReadIni("DM2", "������", App.Path & "\DM2.ini")
                    Else
                        str������ = "0"
                        WriteIni "DM2", "������", "0", App.Path & "\DM2.ini"
                    End If
                    If Not (str������ Like "##" Or str������ Like "#") Then
                        str������ = "0"
                        WriteIni "DM2", "������", "0", App.Path & "\DM2.ini"
                    End If
                    
                    mstr������ = "[" & IIf(Len(str������) = 1, " " & str������, str������)
                End If
                str_SendItem = mstr������ & ",801,01,0000,00,0"
                str_SendItem = str_SendItem & "," & IIf(i_���� = 1, "ST", "RO")             'ST-����,RO-����
                str_SendItem = str_SendItem & ",SE"                                         'Ѫ�� ,SF ��Һ,UR ��(���),
                str_SendItem = str_SendItem & "," & str_���� & Space(15 - Len(str_����))    '�����,15λ,����15λ���ո�
                str_SendItem = str_SendItem & "," & Space(20)
                str_SendItem = str_SendItem & "," & Space(12)
                str_SendItem = str_SendItem & "," & Space(25)
                str_SendItem = str_SendItem & "," & Space(18)
                str_SendItem = str_SendItem & "," & Space(15)
                str_SendItem = str_SendItem & "," & Space(1)
                str_SendItem = str_SendItem & "," & str_�걾�� & Space(15 - Len(str_�걾��)) '�걾��
                str_SendItem = str_SendItem & "," & Space(18)
                str_SendItem = str_SendItem & "," & Space(8)
                str_SendItem = str_SendItem & "," & Space(4)
                str_SendItem = str_SendItem & ",ZLSOFT" & Space(14)                         'λ��
                str_SendItem = str_SendItem & "," & Space(3)                                '����
                str_SendItem = str_SendItem & ",5"                                          '���䵥λ,Ĭ��5
                str_SendItem = str_SendItem & "," & Space(8)                                '����
                str_SendItem = str_SendItem & ",M"                                          '�Ա�
                str_SendItem = str_SendItem & "," & Space(45)
                str_SendItem = str_SendItem & ",111.1  "
                str_SendItem = str_SendItem & ",22.2"
                str_SendItem = str_SendItem & ",33.3"
                str_SendItem = str_SendItem & ",03"
                str_SendItem = str_SendItem & ",4.4444"
                str_SendItem = str_SendItem & "," & str_������Ŀ & "]"
                str_SendItem = str_SendItem & GetCheckCode(str_SendItem) & vbCrLf
                
                If SendSample = "" Then
                    SendSample = Chr(&H4) & Chr(&H1)
                    mstrWaitSend = mstrWaitSend & "|" & str_SendItem & "|" & Chr(&H4)                          '���� ��Ϣ������־(EOT)
                Else
                    mstrWaitSend = mstrWaitSend & "|" & str_SendItem & "|" & Chr(&H4)
                End If
                iSendStep = 1
            End If
        Next
    Else
        str_In = strResponse
        If InStr(str_In, Chr(&H6)) > 0 Or InStr(str_In, Chr(&H3)) > 0 Then
            If mstrWaitSend <> "" Then
                If Left(mstrWaitSend, 1) = "|" Then
                    mstrWaitSend = Mid(mstrWaitSend, 2)
                End If
                If InStr(mstrWaitSend, "|") > 0 Then
                    str_WaitSendLine = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, "|") - 1)
                    mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, "|"))
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


Private Function ACKChar() As String
    If mstrACK = Chr(&H6) Then
        mstrACK = Chr(&H3)
    Else
        mstrACK = Chr(&H6)
    End If
    ACKChar = mstrACK
End Function


Private Function GetCheckCode(ByVal str_In As String) As String
    '����У����
    Dim ls_ch As String
    Dim ll_i As Long
    Dim ll_len As Long
    Dim ll_sum As Long

    If ((IsNull(str_In)) Or (str_In = "")) Then
        GetCheckCode = ""
    End If

    ll_len = Len(str_In)

    For ll_i = 1 To ll_len
        ls_ch = Mid(str_In, ll_i, 1)
        ll_sum = ll_sum + Asc(ls_ch)
    Next

    ll_sum = 256 - (ll_sum Mod 256)
    GetCheckCode = Hex(ll_sum)
    If Len(GetCheckCode) = 1 Then GetCheckCode = "0" & GetCheckCode

    
End Function

