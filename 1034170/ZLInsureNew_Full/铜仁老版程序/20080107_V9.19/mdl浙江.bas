Attribute VB_Name = "mdl�㽭"
Option Explicit
'��ѯ����
Public Declare Function QUERY_HANDLE Lib "SiInterface.DLL" (ByVal InputData As String, ByVal OutputData As String) As Long
'���ײ���
Public Declare Function BUSINESS_HANDLE Lib "SiInterface.DLL" (ByVal InputData As String, ByVal OutputData As String) As Long
'�˹�Ӧ��
Public Declare Function TRADE_ANSWER Lib "SiInterface.DLL" (ByVal InputData As String, ByVal OutputData As String) As Long
'���Ӵ��۽���(intType:1����,2����)
Public Declare Function UF_DLPK Lib "CardOpe.dll" (ByVal intType As Integer, ByRef strPass As String, ByRef dbl��� As Double) As Long
'��ȡ��������(intPathID:1--MF 11,2--MF 12,3--DF04 31,4--DF04 32,5--DF04 33,6--DF04 34,7--DF04 35,8--DF04 36)
Public Declare Function UF_Read_Info Lib "CardOpe.dll" (ByVal intPathID As Integer, ByRef strPass As String, _
    ByRef strInfo As Byte) As Long
'�Կ���ָ�����ݽ����޸�(intPathID:1--MF 11,2--MF 12,3--DF04 31,4--DF04 32,5--DF04 33,6--DF04 34,7--DF04 35,8--DF04 36)
Public Declare Function UF_Update_Info Lib "CardOpe.dll" (ByVal intPathID As Integer, ByRef strPass As String, _
    ByRef strInfo As Byte) As Long
'��ȡ������Ϣ
Public Declare Function GetErrorDesc Lib "CardOpe.dll" (ByRef strDesc As Byte) As Long
Public Declare Function readCardID Lib "cardhandle.DLL" Alias "readCard" (ByRef strCardID As String) As Long

Public gcn�㽭 As New ADODB.Connection, int�㽭���� As Integer, gstrInfo As String

Private str����� As String, mstr���� As String

Public Function CheckReturn�㽭(Optional int���÷�ʽ As Integer = 0) As Boolean
    Dim strDesc As String, bytDesc(2048) As Byte
    If int���÷�ʽ = 1 Then
        If glngReturn < 0 Then
            glngReturn = GetErrorDesc(bytDesc(0))
            strDesc = StrConv(bytDesc, vbUnicode)
            strDesc = Trim(Split(strDesc, Chr(0))(0))
            MsgBox "�ڽ���ҽ������ʱ��ҽ���������´���" & vbCrLf & "    " & strDesc, vbInformation, "�ӿڴ���"
            Exit Function
        End If
    Else
        If glngReturn < 0 Then
            If MsgBox("�ڽ���ҽ������ʱ��������" & vbCrLf & gstrInfo, vbInformation + vbRetryCancel, "�ӿڴ���") = vbRetry Then
                gstrInfo = ""
            Else
                gstrInfo = "-1"
            End If
            Exit Function
        ElseIf IsNumeric(Left(gstrInfo, InStr(gstrInfo, "|") - 1)) Then
            If Val(Left(gstrInfo, InStr(gstrInfo, "|") - 1)) < 0 Then
                If MsgBox("�ڽ���ҽ������ʱ��������" & vbCrLf & gstrInfo, vbInformation + vbRetryCancel, "�ӿڴ���") = vbRetry Then
                    gstrInfo = ""
                Else
                    gstrInfo = "-1"
                End If
                Exit Function
            End If
        Else
            If MsgBox("�ڽ���ҽ������ʱ��������" & vbCrLf & gstrInfo, vbInformation + vbRetryCancel, "�ӿڴ���") = vbRetry Then
                gstrInfo = ""
            Else
                gstrInfo = "-1"
            End If
            Exit Function
        End If
    End If
    gstrInfo = Mid(gstrInfo, InStr(gstrInfo, "|") + 1)
    CheckReturn�㽭 = True
End Function

Public Function Get���ղ���_�㽭(ByVal str������ As String) As String
'���ܣ���ñ��ղ���
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.������,A.����ֵ from ���ղ��� A " & _
              " where A.������='" & str������ & "' and A.����=" & TYPE_�㽭 & " and A.���� is null "
    Call OpenRecordset(rsTemp, gstrSysName)
    
    If rsTemp.EOF = False Then
        Get���ղ���_�㽭 = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
    End If
End Function

Private Function Get����ID(strҽ���� As String) As Long
'���ܣ�ͨ��ҽ�����ĺ����ҽ�����������ID
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select ����ID from �����ʻ� where ���� = '" & gintInsure & "' and ҽ���� = '" & strҽ���� & "'"
    Call OpenRecordset(rsTmp, gstrSysName)
    If Not rsTmp.BOF Then
        Get����ID = CLng(rsTmp("����ID"))
    Else
        Get����ID = 0
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Get����ID = 0
End Function

Private Function Get����(lng����ID As Long, Optional flag As Boolean = False) As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select * From �����ʻ� Where ����=" & gintInsure & " And ����id=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    If Not rsTemp.EOF Then
        If flag Then
            Get���� = Nvl(rsTemp!ҽ����)
        Else
            Get���� = Nvl(rsTemp!����)
        End If
    Else
        Get���� = ""
    End If
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Get���� = ""
End Function

Public Function openConn�㽭() As Boolean
    If gcn�㽭.State = 1 Then
        openConn�㽭 = True
        Exit Function
    End If
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSql As String, rs�㽭 As New ADODB.Recordset
    '��������Ѿ��򿪣��ǾͲ����ٲ���
    If gcn�㽭.State = adStateOpen Then
        openConn�㽭 = True
        Exit Function
    End If
     
    On Error GoTo errH
    
    '���ȶ���������������
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & TYPE_�㽭
    Call OpenRecordset(rsTemp, gstrSysName)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "�㽭������"
                strServer = strTemp
            Case "�㽭�û���"
                strUser = strTemp
            Case "�㽭�û�����"
                strPass = strTemp
            Case "���õ���"
                int�㽭���� = Val(strTemp)
        End Select
        rsTemp.MoveNext
    Loop
    
    On Error Resume Next
    If int�㽭���� = 0 Then
        gcn�㽭.CursorLocation = adUseClient
        gcn�㽭.Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer, strUser, strPass
    End If
    If Err.Number = 0 Then
        openConn�㽭 = True
    Else
        openConn�㽭 = False
    End If
    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ҽ����ʼ��_�㽭() As Boolean
'���ܣ������Ƿ�������ӵ�ǰ�÷�������
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSql As String, rs�㽭 As New ADODB.Recordset
    '��������Ѿ��򿪣��ǾͲ����ٲ���
    If gcn�㽭.State = adStateOpen Then
        gstrSQL = "Select * From ��������Ŀ¼ Where ����=" & gintInsure
        Call OpenRecordset(rsTemp, "ҽ����ʼ��")
        gstrҽ���������� = rsTemp!����
        gstrSQL = "Select * From ������� Where ���=" & gintInsure
        Call OpenRecordset(rsTemp, "ҽ����ʼ��")
        gstrҽԺ���� = Trim(rsTemp!ҽԺ����)
        ҽ����ʼ��_�㽭 = True
        Exit Function
    End If
     
    On Error GoTo errH
    
    '���ȶ���������������
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & TYPE_�㽭
    Call OpenRecordset(rsTemp, gstrSysName)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "�㽭������"
                strServer = strTemp
            Case "�㽭�û���"
                strUser = strTemp
            Case "�㽭�û�����"
                strPass = strTemp
            Case "���õ���"
                int�㽭���� = Val(strTemp)
        End Select
        rsTemp.MoveNext
    Loop
    
    On Error Resume Next
    If int�㽭���� = 0 Then
        gcn�㽭.CursorLocation = adUseClient
        gcn�㽭.Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer, strUser, strPass
    End If
    
    If Err <> 0 Then
        MsgBox "ҽ��ǰ�÷���������ʧ�ܡ�", vbInformation, gstrSysName
        ҽ����ʼ��_�㽭 = False
        Exit Function
    End If
    gstrSQL = "Select * From ��������Ŀ¼ Where ����=" & gintInsure
    Call OpenRecordset(rsTemp, "ҽ����ʼ��")
    gstrҽ���������� = rsTemp!����
    gstrSQL = "Select * From ������� Where ���=" & gintInsure
    Call OpenRecordset(rsTemp, "ҽ����ʼ��")
    gstrҽԺ���� = Trim(rsTemp!ҽԺ����)
    
    ҽ����ʼ��_�㽭 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    ҽ����ʼ��_�㽭 = False
End Function

Public Function �������_�㽭(lng����ID As Long) As Currency
'���ܣ�ͨ�����˵���Ϣ����������
    
    On Error GoTo errHandle
    glngReturn = QUERY_HANDLE("13|" & Get����(lng����ID, True) & "|DF0432|", gstrInfo)
    If CheckReturn�㽭() = False Then
        MsgBox "��ȡ�����ʻ����ʧ��", vbInformation, gstrSysName
        �������_�㽭 = 0
    Else
        �������_�㽭 = Trim(Split(gstrInfo, "|")(0))
    End If
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�㽭 & ",'�ʻ����','" & �������_�㽭 & "')"
    Call ExecuteProcedure("��ݱ�ʶ_�㽭")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    �������_�㽭 = 0
End Function

Public Function ��ݱ�ʶ_�㽭(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim frmIDentified As New frmIdentify�㽭
    Dim strPatiInfo As String, cur��� As Currency
    Dim arr, datCurr As Date, str����� As String
    Dim strSql As String, rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    strPatiInfo = frmIDentified.GetPatient(bytType, mstr����)
    
    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '�������˵�����Ϣ�������ʽ��
        '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����;9.˳���;
        '10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
        '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)
        lng����ID = BuildPatiInfo(bytType, strPatiInfo, lng����ID)
        
        '���ظ�ʽ:�м���벡��ID
        strPatiInfo = frmIDentified.mstrPatient & lng����ID & ";" & frmIDentified.mstrOther
        Unload frmIDentified
    Else
        ��ݱ�ʶ_�㽭 = ""
        MsgBox "ҽ��������Ϣ��ȡʧ��", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    ��ݱ�ʶ_�㽭 = strPatiInfo
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_�㽭 = ""
End Function

Public Function �����������_�㽭(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
'������rsDetail     ������ϸ(����)
'    ����ID         adBigInt, 19, adFldIsNullable
'    �շ����       adVarChar, 2, adFldIsNullable
'    �վݷ�Ŀ       adVarChar, 20, adFldIsNullable
'    ���㵥λ       adVarChar, 6, adFldIsNullable
'    ������         adVarChar, 20, adFldIsNullable
'    �շ�ϸĿID     adBigInt, 19, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ʵ�ս��       adSingle, 15, adFldIsNullable
'    ͳ����       adSingle, 15, adFldIsNullable
'    ����֧������ID adBigInt, 19, adFldIsNullable
'    �Ƿ�ҽ��       adBigInt, 19, adFldIsNullable
'    ժҪ           adVarChar, 200, adFldIsNullable
'    �Ƿ���       adBigInt, 19, adFldIsNullable
'    str���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    Dim lng����ID As Long, rsTemp As New ADODB.Recordset, str���� As String, datCurr As Date, _
        strTemp As String, cur����֧�� As Currency, curͳ��֧�� As Currency, cur������֧�� As Currency, _
        cur����Ա���� As Currency, bytTemp(2048) As Byte
    Dim cur�����ܶ� As Currency
    
    On Error GoTo errHandle
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�з��ã����ܽ���Ԥ���㡣", vbInformation, gstrSysName
        �����������_�㽭 = False
        Exit Function
    End If
    cur�����ܶ� = 0
    While Not rs��ϸ.EOF
        cur�����ܶ� = cur�����ܶ� + rs��ϸ!ʵ�ս��
        rs��ϸ.MoveNext
    Wend
    WriteInfo "��ʼԤ����"
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ("����ID")
    If ������ϸ����_�㽭(0, rs��ϸ) = False Then Exit Function
    
    str���� = Get����(lng����ID)
    datCurr = zlDatabase.Currentdate
    
    strTemp = "09|1|" & UserInfo.���� & "|" & str����� & "|" & str���� & "|" & str����� & "|||" & _
        Format(datCurr, "yyyymmdd|yyyymmdd") & "|0|||||1|" & Trim(gstrҽԺ����) & "|"
    WriteInfo "���ã�" & strTemp
    gstrInfo = Space(1024)
    glngReturn = BUSINESS_HANDLE(strTemp, gstrInfo)
    WriteInfo "���أ�" & gstrInfo
    If CheckReturn�㽭() = False Then Exit Function
    If UBound(Split(gstrInfo, "|")) < 18 Then
        MsgBox "ҽ��Ԥ�������ݸ�ʽ��������ǰ�û���ҽ���������������Ƿ�������", vbInformation, gstrSysName
        Exit Function
    End If
    If cur�����ܶ� <> Val(Split(gstrInfo, "|")(1)) Then
        If MsgBox("ҽ�����ķ��ط����ܶ��뷢���������˶�" & vbCrLf & "����������:" & cur�����ܶ� & "���������ķ���:" & Split(gstrInfo, "|")(1) & vbCrLf & "�Ƿ����ִ�У�", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    cur����֧�� = Val(Split(gstrInfo, "|")(10)) + Val(Split(gstrInfo, "|")(11))
    curͳ��֧�� = Val(Split(gstrInfo, "|")(12))
    cur������֧�� = Val(Split(gstrInfo, "|")(14))
    cur����Ա���� = Val(Split(gstrInfo, "|")(15))
    If cur����֧�� <> 0 Then str���㷽ʽ = "�����ʻ�;" & cur����֧�� & ";0"
    If curͳ��֧�� <> 0 Then str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ <> "", "|", "") & "ͳ�����;" & curͳ��֧�� & ";0"
    If cur������֧�� <> 0 Then str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ <> "", "|", "") & "������֧��;" & cur������֧�� & ";0"
    If cur����Ա���� <> 0 Then str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ <> "", "|", "") & "����Ա����;" & cur����Ա���� & ";0"
    
    �����������_�㽭 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_�㽭(lng����ID As Long, cur����֧�� As Currency, strҽ���� As String, curȫ�Ը� As Currency, cur���Ը� As Currency, curҽ������ As Currency) As Boolean
'���ܣ���������ý�����ϸ���ݲ��ҽ��н���
'������������ϸ����ʧ�ܣ���ֱ�ӽ������������غ���ʧ��
    Dim strSql As String, rsTmp As New ADODB.Recordset
    Dim rs�㽭 As New ADODB.Recordset, lng����ID As Long, strTemp As String
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, datCurr As Date
    Dim cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim cur�����ʻ� As Currency, cur��� As Currency, cur�������� As Currency, curͳ��֧�� As Currency
    Dim cur������֧�� As Currency, cur����Ա���� As Currency, str������ˮ�� As String, str���� As String
    On Error GoTo errHandle
    '����������㣬�޷����н���
    lng����ID = Get����ID(strҽ����)
    cur��� = �������_�㽭(lng����ID)
    
    str���� = Get����(lng����ID)
    datCurr = zlDatabase.Currentdate
    WriteInfo "��ʼ����"
    strTemp = "10|1|" & UserInfo.���� & "|" & str����� & "|" & str���� & "|" & str����� & "|||" & _
        Format(datCurr, "yyyymmdd|yyyymmdd") & "|0|||||0|" & Trim(gstrҽԺ����) & "|"
    WriteInfo "���ã�" & strTemp
    glngReturn = BUSINESS_HANDLE(strTemp, gstrInfo)
    WriteInfo "���أ�" & gstrInfo
    If CheckReturn�㽭() = False Then Exit Function
    
    str������ˮ�� = Split(gstrInfo, "|")(0)
    cur�������� = Val(Split(gstrInfo, "|")(1))
    cur����֧�� = Val(Split(gstrInfo, "|")(10)) + Val(Split(gstrInfo, "|")(11))
    curͳ��֧�� = Val(Split(gstrInfo, "|")(12))
    cur������֧�� = Val(Split(gstrInfo, "|")(14))
    cur����Ա���� = Val(Split(gstrInfo, "|")(15))
    
    WriteInfo "Ӧ��" & str������ˮ��
    glngReturn = TRADE_ANSWER(str������ˮ��, gstrInfo)
    WriteInfo "���أ�" & gstrInfo
    If CheckReturn�㽭() = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & _
            "," & gintInsure & "," & Year(datCurr) & "," & cur�ʻ������ۼ� & _
            "," & cur�ʻ�֧���ۼ� + cur����֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",null,null,null,null)"
    Call ExecuteProcedure("�㽭ҽ��")
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & _
            lng����ID & "," & Year(datCurr) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� + cur����֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� + curͳ��֧�� + cur������֧�� + cur����Ա���� & "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & _
            cur�������� & "," & curȫ�Ը� & "," & cur���Ը� & ",NULL," & curͳ��֧�� + cur������֧�� + cur����Ա���� & ",NULL,NULL," & _
            cur����֧�� & ",NULL,NULL,NULL,'" & str������ˮ�� & "')"
    Call ExecuteProcedure("�㽭ҽ��")
    
    cur��� = �������_�㽭(lng����ID)
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�㽭 & ",'�ʻ����','" & cur��� & "')"
    Call ExecuteProcedure("�㽭ҽ��")
    
    �������_�㽭 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    �������_�㽭 = False
End Function

Public Function ����������_�㽭(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, str������ As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, lng����ID As Long, strTemp As String
    Dim cur��� As Currency, cur�������� As Currency, curͳ��֧�� As Currency
    Dim cur������֧�� As Currency, cur����Ա���� As Currency, str������ˮ�� As String
    Dim datCurr As Date

    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B" & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    lng����ID = rsTemp("����ID")
    WriteInfo "׼������"
    'ȡԭ���ݽ�����ˮ��
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=" & gintInsure & " and ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    If rsTemp.EOF = True Then
        MsgBox "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If IsNull(rsTemp!��ע) Then
        MsgBox "�õ��ݵĽ�����ˮ�Ŷ�ʧ���������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    str������ = rsTemp!��ע
    strTemp = "99|" & str������ & "|" & Trim(gstrҽԺ����) & "|"
    WriteInfo "���ã�" & strTemp
    '���ýӿ�������
    glngReturn = BUSINESS_HANDLE(strTemp, gstrInfo)
    WriteInfo "���أ�" & gstrInfo
    If CheckReturn�㽭() = False Then Exit Function
    
    str������ = Split(gstrInfo, "|")(0)
    cur�������� = Val(Split(gstrInfo, "|")(1))
    cur�����ʻ� = Val(Split(gstrInfo, "|")(10)) + Val(Split(gstrInfo, "|")(11))
    curͳ��֧�� = Val(Split(gstrInfo, "|")(12))
    cur������֧�� = Val(Split(gstrInfo, "|")(14))
    cur����Ա���� = Val(Split(gstrInfo, "|")(15))
    
'    WriteInfo "Ӧ��" & str������
'    glngReturn = TRADE_ANSWER(str������, gstrInfo)
'    WriteInfo "���أ�" & gstrInfo
'    If CheckReturn�㽭() = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & _
            lng����ID & "," & Year(datCurr) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� - curͳ��֧�� - cur������֧�� - cur����Ա���� & "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & _
            0 - cur�������� & ",0,0,NULL," & 0 - (curͳ��֧�� + cur������֧�� + cur����Ա����) & ",NULL,NULL," & _
            0 - cur�����ʻ� & ",NULL,NULL,NULL,'" & str������ & "')"
    Call ExecuteProcedure("�㽭ҽ��")
    
    cur��� = �������_�㽭(lng����ID)
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�㽭 & ",'�ʻ����','" & cur��� & "')"
    Call ExecuteProcedure("�㽭ҽ��")
    
    ����������_�㽭 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ������ϸ����_�㽭(lng����ID As Long, Optional rs��ϸIN As ADODB.Recordset = Nothing, Optional strסԺ�� As String = "") As Boolean
    Dim lng����ID  As Long, rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset, _
        str����Ա As String, cur��������, str������ As String, datCurr As Date, rs���� As New ADODB.Recordset, _
        strTemp As String, iLoop As Long, str��ϸ���� As String, str��ϸ���� As String, str��ϸ���� As String, _
        str�շ���� As String, strҩƷ�ȼ� As String, str������ As String, cur�Ը����� As Currency
'    ����ID         adBigInt, 19, adFldIsNullable
'    �շ����       adVarChar, 2, adFldIsNullable
'    �վݷ�Ŀ       adVarChar, 20, adFldIsNullable
'    ���㵥λ       adVarChar, 6, adFldIsNullable
'    ������         adVarChar, 20, adFldIsNullable
'    �շ�ϸĿID     adBigInt, 19, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ʵ�ս��       adSingle, 15, adFldIsNullable
'    ͳ����       adSingle, 15, adFldIsNullable
'    ����֧������ID adBigInt, 19, adFldIsNullable
'    �Ƿ�ҽ��       adBigInt, 19, adFldIsNullable
'    ժҪ           adVarChar, 200, adFldIsNullable
'    �Ƿ���       adBigInt, 19, adFldIsNullable
'    str���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    WriteInfo vbCrLf & "��ʼ������ϸ����"
    If rs��ϸIN Is Nothing Then
        gstrSQL = "Select * From ���˷��ü�¼ Where ��¼״̬<>0 And Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 and ����ID=" & lng����ID
        Call OpenRecordset(rs��ϸ, gstrSysName)
    Else
        Set rs��ϸ = rs��ϸIN.Clone
    End If
    
    If rs��ϸ.EOF = True Then
        ������ϸ����_�㽭 = False
        Exit Function
    End If
    
    lng����ID = rs��ϸ!����ID
    str����Ա = ToVarchar(UserInfo.����, 20)
    Randomize
    If strסԺ�� = "" Then
        str����� = Chr(Year(Date) - 1939) & Hex(Month(datCurr)) & IIf(Day(datCurr) < 10, Day(datCurr), Chr(Day(datCurr) + 55)) & Format(datCurr, "HHMMSS") & Format(999 * Rnd + 1, "0##")
        strסԺ�� = str�����
    Else
        str����� = Chr(Year(Date) - 1939) & Hex(Month(datCurr)) & IIf(Day(datCurr) < 10, Day(datCurr), Chr(Day(datCurr) + 55)) & Format(datCurr, "HHMMSS") & Format(999 * Rnd + 1, "0##")
    End If
    str������ = Format(datCurr, "yyyymmddHHMMSS") & Format(999 * Rnd + 1, "0##") & Format(gstrҽԺ����, "0####") & "00000110"
    
    iLoop = 1
    'д������ϸ
    Do Until rs��ϸ.EOF
        gstrSQL = "Select * From �շ�ϸĿ Where ID=" & rs��ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, gstrSysName)
        str��ϸ���� = rsTemp!����
        str��ϸ���� = rsTemp!����
        Select Case rsTemp!���
            Case "5"
                str�շ���� = "11"
            Case "6"
                str�շ���� = "12"
            Case "7"
                str�շ���� = "13"
            Case "C"
                str�շ���� = "25"
            Case "D"
                str�շ���� = "21"
            Case "E"
                str�շ���� = "31"
            Case "F"
                str�շ���� = "24"
            Case "G"
                str�շ���� = "91"
            Case "H"
                str�շ���� = "33"
            Case "I"
                str�շ���� = "91"
            Case "J"
                str�շ���� = "34"
            Case "K"
                str�շ���� = "26"
            Case "L"
                str�շ���� = "23"
            Case "M"
                str�շ���� = "91"
            Case "Z"
                str�շ���� = "91"
            Case "1"
                str�շ���� = "91"
            Case Else
                str�շ���� = "91"
        End Select
        gstrSQL = "Select * From ����֧����Ŀ Where ����=" & gintInsure & " And �Ƿ�ҽ��=1 And �շ�ϸĿID=" & rs��ϸ!�շ�ϸĿID
'        gstrSQL = "Select A.*,B.���� As ��������,B.���� As ������� from ����֧����Ŀ A,����֧������ B Where A.����=B.���� And " & _
            "A.����ID=B.ID And A.����=" & gintInsure & " And A.�Ƿ�ҽ��=1 And A.�շ�ϸĿID=" & rs��ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, gstrSysName)
        If rsTemp.EOF Then
            str��ϸ���� = str��ϸ����
            If InStr("11 12", str�շ����) > 0 Then
                strҩƷ�ȼ� = "3"
            Else
                strҩƷ�ȼ� = "1"
            End If
        Else
            If InStr("11 12", str�շ����) > 0 Then
                Select Case Nvl(rsTemp!��ע, "")
                    Case "����"
                        strҩƷ�ȼ� = "1"
                    Case "����"
                        strҩƷ�ȼ� = "2"
                    Case Else
                        strҩƷ�ȼ� = "3"
                End Select
            Else
                Select Case Nvl(rsTemp!��ע, "")
                    Case "����"
                        strҩƷ�ȼ� = "3"
                    Case "����"
                        strҩƷ�ȼ� = "2"
                    Case Else
                        strҩƷ�ȼ� = "1"
                End Select
                If Nvl(rsTemp!����id, "") <> "" Then
                    gstrSQL = "Select * From ����֧������ Where ID=" & rsTemp!����id & " And ����=" & gintInsure
                    Call OpenRecordset(rs����, gstrSysName)
                    If Not rs����.EOF Then
                        If Nvl(rs����!����, "") = "����" Or Nvl(rs����!����, "") = "������" Or Nvl(rs����!����, "") = "��������" Then
                            str�շ���� = "22"
                        End If
                    End If
                End If
            End If
            str��ϸ���� = Nvl(rsTemp!��Ŀ����, str��ϸ����)
'            str��ϸ���� = Nvl(rsTemp!��Ŀ����, str��ϸ����)
        End If
        
        If strҩƷ�ȼ� = "2" Then
            If str�շ���� = "11" Or str�շ���� = "12" Or str�շ���� = "13" Then
                strTemp = "Select AKA069 From KA02 Where AKA060='" & str��ϸ���� & "'"
            Else
                strTemp = "Select AKA069 From KA03 Where AKA090='" & str��ϸ���� & "'"
            End If
            Set rsTemp = gcn�㽭.Execute(strTemp)
            If rsTemp.EOF Then
                cur�Ը����� = 0.05
            Else
                cur�Ը����� = rsTemp(0)
            End If
        ElseIf strҩƷ�ȼ� = "3" Then
            cur�Ը����� = 1
        Else
            cur�Ը����� = 0
        End If
        
        str������ = 0
        If strҩƷ�ȼ� = "1" Then
            str������ = 0
        ElseIf strҩƷ�ȼ� = "2" Then
            If rs��ϸ.Fields.Count < 26 Then
                str������ = rs��ϸ!ʵ�ս�� * cur�Ը�����
            Else
                str������ = rs��ϸ!��� * cur�Ը�����
            End If
        ElseIf strҩƷ�ȼ� = "3" Then
            If rs��ϸ.Fields.Count < 26 Then
                str������ = rs��ϸ!ʵ�ս��
            Else
                str������ = rs��ϸ!���
            End If
        End If
        If rs��ϸ.Fields.Count < 26 Then
            strTemp = "Insert Into KC22 (AKB020,AKC190,CKC250,AAE072,CKC130,CKC131,AKC515,AKC221,AKA063,AKC220," & _
                "AKC222,AKC223,AKC224,AKA065,AKC225,AKC226,AKC227,AKC228,AKC253) Values ('" & Trim(gstrҽԺ����) & "','" & _
                strסԺ�� & "','" & iLoop & "','" & str����� & "',1,'" & str������ & "','" & str��ϸ���� & "',to_date('" & _
                Format(datCurr, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'),'" & str�շ���� & "','" & _
                str����� & "','" & str��ϸ���� & "','" & str��ϸ���� & "','" & IIf(InStr("11 12", str�շ����) > 0, "0", IIf(str�շ���� = "13", "1", "2")) & "','" & _
                strҩƷ�ȼ� & "'," & rs��ϸ!���� & "," & rs��ϸ!���� & "," & rs��ϸ!ʵ�ս�� & "," & IIf(strҩƷ�ȼ� = "3", "0," & str������, str������ & ",0") & ")"
        Else
            gstrSQL = "Select * From ���˷��ü�¼ Where NO='" & rs��ϸ!NO & "' And ���=" & rs��ϸ!��� & " And �����־=2 And ��¼����=" & rs��ϸ!��¼���� & " And ��¼״̬=" & rs��ϸ!��¼״̬
            Call OpenRecordset(rsTemp, gstrSysName)
            strTemp = "Insert Into KC22 (AKB020,AKC190,CKC250,AAE072,CKC130,CKC131,AKC515,AKC221,AKA063,AKC220," & _
                "AKC222,AKC223,AKC224,AKA065,AKC225,AKC226,AKC227,AKC228,AKC253) Values ('" & Trim(gstrҽԺ����) & "','" & _
                strסԺ�� & "','" & iLoop & "','" & str����� & "',1,'" & str������ & "','" & str��ϸ���� & "',to_date('" & _
                Format(rsTemp!����ʱ��, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'),'" & str�շ���� & "','" & _
                str����� & "','" & str��ϸ���� & "','" & str��ϸ���� & "','" & IIf(InStr("11 12", str�շ����) > 0, "0", IIf(str�շ���� = "13", "1", "2")) & "','" & _
                strҩƷ�ȼ� & "'," & rs��ϸ!�۸� & "," & rs��ϸ!���� & "," & rs��ϸ!��� & "," & IIf(strҩƷ�ȼ� = "3", "0," & str������, str������ & ",0") & ")"
        End If
        WriteInfo strTemp
        If rs��ϸ.Fields.Count >= 26 Then
            If Nvl(rsTemp!�Ƿ��ϴ�, 0) = 0 Then gcn�㽭.Execute strTemp
            gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rsTemp("ID") & "')"
            Call ExecuteProcedure(gstrSysName)
        Else
            gcn�㽭.Execute strTemp
        End If
        rs��ϸ.MoveNext
        iLoop = iLoop + 1
    Loop
    ������ϸ����_�㽭 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_�㽭(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset, str���� As String, datCurr As Date, strInNote As String

    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,D.���� as ���ұ���,A.��Ժ����,A.סԺҽʦ,C.����," & _
            "C.���� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = " & lng��ҳID & " And A.����ID = " & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "δ�ܻ�ȡ��Ժ���˵������Ϣ��", vbInformation, gstrSysName
        ��Ժ�Ǽ�_�㽭 = False
        Exit Function
    End If

    '��ȡ��Ժ��ϣ����ֱ��룩
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID, True, True, True) '��Ժ���
    If strInNote <> "" Then
        strInNote = Mid(strInNote, InStr(strInNote, "|") + 1)
    End If
    
    str���� = Get����(lng����ID)
    WriteInfo "������Ժ�Ǽ�"
    WriteInfo "���ã�" & "01|" & str���� & "|1|ZY" & lng����ID & "_" & lng��ҳID & "|" & lng����ID & "_" & lng��ҳID & "|" & _
        UserInfo.��� & "|0|0||2|" & Format(datCurr, "yyyymmdd") & "|0|" & strInNote & "|" & Trim(gstrҽԺ����) & "|"
    '���ýӿ�������
    glngReturn = BUSINESS_HANDLE("01|" & str���� & "|1|ZY" & lng����ID & "_" & lng��ҳID & "|" & lng����ID & "_" & lng��ҳID & "|" & _
        UserInfo.��� & "|0|0||2|" & Format(datCurr, "yyyymmdd") & "|0|" & strInNote & "|" & Trim(gstrҽԺ����) & "|", gstrInfo)
    WriteInfo "���أ�" & gstrInfo
    If CheckReturn�㽭() = False Then Exit Function
    WriteInfo "�����Ժ�Ǽ�"
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    ��Ժ�Ǽ�_�㽭 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_�㽭 = False
End Function

Public Function ��Ժ�Ǽǳ���_�㽭(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ�Ǽǳ�����Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset, str���� As String, datCurr As Date, strInNote As String

    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,D.���� as ���ұ���,A.��Ժ����,A.סԺҽʦ,C.����," & _
            "C.���� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = " & lng��ҳID & " And A.����ID = " & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "δ�ܻ�ȡ��Ժ���˵������Ϣ��", vbInformation, gstrSysName
        ��Ժ�Ǽǳ���_�㽭 = False
        Exit Function
    End If

    '��ȡ��Ժ��ϣ����ֱ��룩
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID, True, True, True) '��Ժ���
    If strInNote <> "" Then
        strInNote = Mid(strInNote, InStr(strInNote, "|") + 1)
    End If
    
    str���� = Get����(lng����ID)
    WriteInfo "���ã�" & "01|" & str���� & "|-1|ZY" & lng����ID & "_" & lng��ҳID & "|" & lng����ID & "_" & lng��ҳID & "|" & _
        UserInfo.��� & "|0|0||2|" & Format(datCurr, "yyyymmdd") & "|0|" & strInNote & "|" & Trim(gstrҽԺ����) & "|"
    '���ýӿ�������
    glngReturn = BUSINESS_HANDLE("01|" & str���� & "|-1|ZY" & lng����ID & "_" & lng��ҳID & "|" & lng����ID & "_" & lng��ҳID & "|" & _
        UserInfo.��� & "|0|0||2|" & Format(datCurr, "yyyymmdd") & "|0|" & strInNote & "|" & Trim(gstrҽԺ����) & "|", gstrInfo)
    WriteInfo "���أ�" & gstrInfo
    If CheckReturn�㽭() = False Then Exit Function

     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    ��Ժ�Ǽǳ���_�㽭 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽǳ���_�㽭 = False
End Function

Public Function סԺ�������_�㽭(rs��ϸ As Recordset, lng����ID As Long, strҽ���� As String) As String
'������rsDetail     ������ϸ(����)
'    ����ID         adBigInt, 19, adFldIsNullable
'    �շ����       adVarChar, 2, adFldIsNullable
'    �վݷ�Ŀ       adVarChar, 20, adFldIsNullable
'    ���㵥λ       adVarChar, 6, adFldIsNullable
'    ������         adVarChar, 20, adFldIsNullable
'    �շ�ϸĿID     adBigInt, 19, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ʵ�ս��       adSingle, 15, adFldIsNullable
'    ͳ����       adSingle, 15, adFldIsNullable
'    ����֧������ID adBigInt, 19, adFldIsNullable
'    �Ƿ�ҽ��       adBigInt, 19, adFldIsNullable
'    ժҪ           adVarChar, 200, adFldIsNullable
'    �Ƿ���       adBigInt, 19, adFldIsNullable
'    str���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    Dim rsTemp As New ADODB.Recordset, str���� As String, datCurr As Date, _
        strTemp As String, cur����֧�� As Currency, curͳ��֧�� As Currency, cur������֧�� As Currency, _
        cur����Ա���� As Currency, bytTemp(2048) As Byte, lng��ҳID As Long
    Dim cur�����ܶ� As Currency, strסԺ�� As String, str���㷽ʽ As String
    
    On Error GoTo errHandle
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�з��ã����ܽ���Ԥ���㡣", vbInformation, gstrSysName
        Exit Function
    End If
    cur�����ܶ� = 0
    While Not rs��ϸ.EOF
        cur�����ܶ� = cur�����ܶ� + rs��ϸ!���
        rs��ϸ.MoveNext
    Wend
    WriteInfo "��ʼԤ����"
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ!����ID
    gstrSQL = "Select max(��ҳid) from ������ҳ Where ����id=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    lng��ҳID = rsTemp(0)
    
    strסԺ�� = lng����ID & "_" & lng��ҳID
    If ������ϸ����_�㽭(0, rs��ϸ, strסԺ��) = False Then Exit Function
    
    str���� = Get����(lng����ID)
    datCurr = zlDatabase.Currentdate
    
    strTemp = "09|2|" & UserInfo.���� & "|" & strסԺ�� & "|" & str���� & "|" & strסԺ�� & "|||" & _
        Format(datCurr, "yyyymmdd|yyyymmdd") & "|0|||||1|" & Trim(gstrҽԺ����) & "|"
    WriteInfo "���ã�" & strTemp
    gstrInfo = Space(1024)
    glngReturn = BUSINESS_HANDLE(strTemp, gstrInfo)
    WriteInfo "���أ�" & gstrInfo
    If CheckReturn�㽭() = False Then Exit Function
    
    If cur�����ܶ� <> Val(Split(gstrInfo, "|")(1)) Then
        If MsgBox("ҽ�����ķ��ط����ܶ��뷢���������˶�" & vbCrLf & "����������:" & cur�����ܶ� & "���������ķ���:" & Split(gstrInfo, "|")(1) & vbCrLf & "�Ƿ����ִ�У�", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    cur����֧�� = Val(Split(gstrInfo, "|")(13)) + Val(Split(gstrInfo, "|")(14))
    curͳ��֧�� = Val(Split(gstrInfo, "|")(15))
    cur������֧�� = Val(Split(gstrInfo, "|")(17))
    cur����Ա���� = Val(Split(gstrInfo, "|")(18))
    If cur����֧�� <> 0 Then str���㷽ʽ = "�����ʻ�;" & cur����֧�� & ";0"
    If curͳ��֧�� <> 0 Then str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ <> "", "|", "") & "ͳ�����;" & curͳ��֧�� & ";0"
    If cur������֧�� <> 0 Then str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ <> "", "|", "") & "������֧��;" & cur������֧�� & ";0"
    If cur����Ա���� <> 0 Then str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ <> "", "|", "") & "����Ա����;" & cur����Ա���� & ";0"
    
    סԺ�������_�㽭 = str���㷽ʽ
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_�㽭(lng����ID As Long, lng����ID As Long) As Boolean
'���ܣ���סԺ���ý�����ϸ���ݲ��ҽ��н���
'���סԺ������ϸ����ʧ�ܣ���ֱ�ӽ������������غ���ʧ��
    Dim strSql As String, rsTemp As New ADODB.Recordset
    Dim rs�㽭 As New ADODB.Recordset, strTemp As String, lng��ҳID As Long, strסԺ�� As String
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, datCurr As Date
    Dim cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim cur����֧�� As Currency, cur��� As Currency, cur�������� As Currency, curͳ��֧�� As Currency
    Dim cur������֧�� As Currency, cur����Ա���� As Currency, str������ˮ�� As String, str���� As String
    On Error GoTo errHandle
    '����������㣬�޷����н���
    gstrSQL = "Select max(��ҳid) from ������ҳ where ����id=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    cur��� = �������_�㽭(lng����ID)
    lng��ҳID = rsTemp(0)
    
    str���� = Get����(lng����ID)
    datCurr = zlDatabase.Currentdate
    WriteInfo "��ʼ����"
    strסԺ�� = lng����ID & "_" & lng��ҳID
    strTemp = "10|2|" & UserInfo.���� & "|" & strסԺ�� & "|" & str���� & "|" & strסԺ�� & "|||" & _
        Format(datCurr, "yyyymmdd|yyyymmdd") & "|0|||||0|" & Trim(gstrҽԺ����) & "|"
    WriteInfo "���ã�" & strTemp
    glngReturn = BUSINESS_HANDLE(strTemp, gstrInfo)
    WriteInfo "���أ�" & gstrInfo
    If CheckReturn�㽭() = False Then Exit Function
    
    str������ˮ�� = Split(gstrInfo, "|")(0)
    cur�������� = Val(Split(gstrInfo, "|")(1))
    cur����֧�� = Val(Split(gstrInfo, "|")(13)) + Val(Split(gstrInfo, "|")(14))
    curͳ��֧�� = Val(Split(gstrInfo, "|")(15))
    cur������֧�� = Val(Split(gstrInfo, "|")(17))
    cur����Ա���� = Val(Split(gstrInfo, "|")(18))
    
    WriteInfo "Ӧ��" & str������ˮ��
    glngReturn = TRADE_ANSWER(str������ˮ��, gstrInfo)
    WriteInfo "���أ�" & gstrInfo
    If CheckReturn�㽭() = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & _
            "," & gintInsure & "," & Year(datCurr) & "," & cur�ʻ������ۼ� & _
            "," & cur�ʻ�֧���ۼ� + cur����֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",null,null,null,null)"
    Call ExecuteProcedure("�㽭ҽ��")
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & _
            lng����ID & "," & Year(datCurr) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� + cur����֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� + curͳ��֧�� + cur������֧�� + cur����Ա���� & "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & _
            cur�������� & ",0,0,NULL," & curͳ��֧�� + cur������֧�� + cur����Ա���� & ",NULL,NULL," & _
            cur����֧�� & ",NULL,NULL,NULL,'" & str������ˮ�� & "')"
    Call ExecuteProcedure("�㽭ҽ��")
    
    cur��� = �������_�㽭(lng����ID)
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�㽭 & ",'�ʻ����','" & cur��� & "')"
    Call ExecuteProcedure("�㽭ҽ��")
    
    סԺ����_�㽭 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    סԺ����_�㽭 = False
End Function

Public Function סԺ�������_�㽭(lng����ID As Long) As Boolean
'���ܣ���סԺ�շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, str������ As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, lng����ID As Long, strTemp As String
    Dim cur��� As Currency, cur�������� As Currency, curͳ��֧�� As Currency
    Dim cur������֧�� As Currency, cur����Ա���� As Currency, str������ˮ�� As String
    Dim datCurr As Date, cur�����ʻ� As Currency, lng����ID As Long

    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    MsgBox "ҽ�����˲��ܽ���סԺ�������", vbInformation, gstrSysName
    Exit Function
    '�˷�
    gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B" & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    lng����ID = rsTemp("����ID")
    WriteInfo "׼������"
    'ȡԭ���ݽ�����ˮ��
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=" & gintInsure & " and ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    If rsTemp.EOF = True Then
        MsgBox "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If IsNull(rsTemp!��ע) Then
        MsgBox "�õ��ݵĽ�����ˮ�Ŷ�ʧ���������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    str������ = rsTemp!��ע
    WriteInfo "���ã�" & "99|" & str������ & "|" & Trim(gstrҽԺ����) & "|"
    '���ýӿ�������
    glngReturn = BUSINESS_HANDLE("99|" & str������ & "|" & Trim(gstrҽԺ����) & "|", gstrInfo)
    WriteInfo "���أ�" & gstrInfo
    If CheckReturn�㽭() = False Then Exit Function
    
    str������ = Split(gstrInfo, "|")(0)
    cur�������� = Val(Split(gstrInfo, "|")(1))
    cur�����ʻ� = Val(Split(gstrInfo, "|")(13)) + Val(Split(gstrInfo, "|")(14))
    curͳ��֧�� = Val(Split(gstrInfo, "|")(15))
    cur������֧�� = Val(Split(gstrInfo, "|")(17))
    cur����Ա���� = Val(Split(gstrInfo, "|")(18))
    
'    WriteInfo "Ӧ��" & str������
'    glngReturn = TRADE_ANSWER(str������, gstrInfo)
'    WriteInfo "���أ�" & gstrInfo
'    If CheckReturn�㽭() = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & _
            lng����ID & "," & Year(datCurr) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� - curͳ��֧�� - cur������֧�� - cur����Ա���� & "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & _
            0 - cur�������� & ",0,0,NULL," & 0 - (curͳ��֧�� + cur������֧�� + cur����Ա����) & ",NULL,NULL," & _
            0 - cur�����ʻ� & ",NULL,NULL,NULL,'" & str������ & "')"
    Call ExecuteProcedure("�㽭ҽ��")
    
    cur��� = �������_�㽭(lng����ID)
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�㽭 & ",'�ʻ����','" & cur��� & "')"
    Call ExecuteProcedure("�㽭ҽ��")
    
    סԺ�������_�㽭 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_�㽭(lng����ID As Long, lng��ҳID As Long) As Boolean
    On Error GoTo errHandle
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    ��Ժ�Ǽǳ���_�㽭 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    ��Ժ�Ǽǳ���_�㽭 = False
End Function

Public Function ��Ժ�Ǽ�_�㽭(lng����ID As Long, lng��ҳID As Long) As Boolean
    On Error GoTo errHandle
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    ��Ժ�Ǽ�_�㽭 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    ��Ժ�Ǽ�_�㽭 = False
End Function
