Attribute VB_Name = "mdl��ɽ"
Option Explicit
Public gcn��ɽ As New ADODB.Connection
Private mstr����� As String

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTTOPMOST = -2

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const GMEM_FIXED = &H0
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Function ҽ����ʼ��_��ɽ() As Boolean
'���ܣ������Ƿ�������ӵ�ǰ�÷�������
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSql As String, rs��ɽ As New ADODB.Recordset
    '��������Ѿ��򿪣��ǾͲ����ٲ���
    If gcn��ɽ.State = adStateOpen Then
        ҽ����ʼ��_��ɽ = True
        Exit Function
    End If
     
    On Error GoTo errH
    
    '���ȶ���������������
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & TYPE_�����ɽ
    Call OpenRecordset(rsTemp, "��ɽҽ��")
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "��ɽ������"
                strServer = strTemp
            Case "��ɽ�û���"
                strUser = strTemp
            Case "��ɽ�û�����"
                strPass = strTemp
        End Select
        rsTemp.MoveNext
    Loop
    
    On Error Resume Next
    If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
        gcn��ɽ.Open "Provider=SQLOLEDB.1;Initial Catalog=hw_interface;Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
    Else
        gcn��ɽ.Open "Driver={Microsoft ODBC for Oracle};Server=" & _
            strServer, strUser, strPass
    End If
    If Err <> 0 Then
        MsgBox "ҽ��ǰ�÷���������ʧ�ܡ�", vbInformation, gstrSysName
        ҽ����ʼ��_��ɽ = False
        Exit Function
    End If
    ҽ����ʼ��_��ɽ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    ҽ����ʼ��_��ɽ = False
End Function

Public Function ��ݱ�ʶ_��ɽ(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim frmIDentified As New frmIdentify��ɽ
    Dim strPatiInfo As String, cur��� As Currency
    Dim arr, datCurr As Date, str����� As String
    Dim strSql As String, str���ⲡ As String
    Dim strTemp As String, errLine As Integer
    
    '�ж��Ƿ񱣴���IC����֤��
    strTemp = Get���ղ���_��ɽ("����֤��")
    If strTemp = "" Then
        MsgBox "����ҽ�����������ñ���ҽ����IC����֤�롣", vbInformation, gstrSysName
        Exit Function
    End If
    
    frmIDentified.mstr��֤�� = strTemp
    frmIDentified.Tag = bytType
    frmIDentified.Show 1
    'New:0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����)
    On Error GoTo errHandle
    strPatiInfo = frmIDentified.mstrPatiInfo: errLine = 1
    cur��� = frmIDentified.mcur���: errLine = 2
    Unload frmIDentified: errLine = 3
    
    If strPatiInfo <> "" Then
        '�������˵�����Ϣ�������ʽ��
        '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����;9.˳���;
        '10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
        '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)

        lng����ID = BuildPatiInfo(bytType, strPatiInfo & ";;;;" & cur��� & ";;;;;;;" & cur��� & ";;;;;", lng����ID): errLine = 4
        '���ظ�ʽ:�м���벡��ID
        strPatiInfo = strPatiInfo & ";" & lng����ID & ";;;;" & cur��� & ";;;;;;;" & cur��� & ";;;;;": errLine = 5
    Else
        ��ݱ�ʶ_��ɽ = "": errLine = 6
        MsgBox "ҽ��������Ϣ��ȡʧ��", vbInformation, gstrSysName
        Exit Function
    End If
    arr = Split(strPatiInfo, ";"): errLine = 12
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
        '����Ƿ����ⲡ
        str���ⲡ = frmIDentified.mstr���ⲡ: errLine = 7
        gstr���ⲡ�� = str���ⲡ: errLine = 8
    Else
        str���ⲡ = Get����ID(CStr(arr(1)), CStr(gintInsure)): errLine = 9
    End If
    If bytType <> 0 Then
        ��ݱ�ʶ_��ɽ = strPatiInfo: errLine = 10
    End If
    '���Ϊ���ﲡ�ˣ��ͽ��Ž�������Ǽ�
    datCurr = zlDatabase.Currentdate: errLine = 11
    str����� = ToVarchar(lng����ID & Format(datCurr, "yyddhhmmss"), 16): errLine = 13
    mstr����� = str�����: errLine = 14
    '��������Ǽ�׼��
    If bytType <> 0 Then
        ��ݱ�ʶ_��ɽ = strPatiInfo
    Else
        strSql = "insert into Check_doex_interface(Bill_no,App_code" & _
                ",Doct_flag,Doex_no,Ill_type,Ic_id,Is_bala,Regi_op_id) values('" & _
                str����� & "','" & Mid(gstrҽԺ����, 1, 4) & "','" & IIf(bytType = 1, 1, 0) & "','" & _
                Left(str�����, 10) & "','" & str���ⲡ & _
                "','" & arr(2) & arr(0) & "','0','" & ToVarchar(UserInfo.����, 8) & "')": errLine = 15
        gcn��ɽ.Execute strSql: errLine = 16
        '��������Ǽ�����
        strSql = "insert into Check_bill_request(Bill_no,App_code," & _
                "Request_status) values('" & str����� & "','" & _
                Mid(gstrҽԺ����, 1, 4) & "','0')": errLine = 17
        gcn��ɽ.Execute strSql: errLine = 18
        If Checkrequest(str�����) = False Then
            'ɾ��ʧ�ܵ�����Ǽǵ�
            strSql = "delete from Check_bill_request where Bill_no = '" & str����� & _
                    "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'": errLine = 19
            gcn��ɽ.Execute strSql: errLine = 10
            strSql = "delete from Check_doex_interface where Bill_no = '" & _
                    str����� & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'": errLine = 21
            gcn��ɽ.Execute strSql: errLine = 22
            ��ݱ�ʶ_��ɽ = ""
            Exit Function
        Else
            ��ݱ�ʶ_��ɽ = strPatiInfo
        End If
    End If
    Exit Function
errHandle:
    MsgBox "���������[�����֤]ģ���" & errLine & "��", vbInformation, "����"
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_��ɽ = ""
End Function

Public Function �������_��ɽ(lng����ID As Long, cur����֧�� As Currency, strҽ���� As String, curȫ�Ը� As Currency, cur���Ը� As Currency, curҽ������ As Currency) As Boolean
'���ܣ���������ý�����ϸ���ݲ��ҽ��н���
'������������ϸ����ʧ�ܣ���ֱ�ӽ������������غ���ʧ��
    Dim strSql As String, rsTmp As New ADODB.Recordset
    Dim rs��ɽ As New ADODB.Recordset
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, curDate As Date
    Dim cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim cur�����ʻ� As Currency, cur���� As Currency, cur����ͳ���޶� As Currency
    Dim cur���ͳ���޶� As Currency, cur�����Ը� As Currency, cur��� As Currency
    Dim cur�������� As Currency, cur�ز�ͳ�� As Currency, str�������� As String
    
    On Error GoTo errHandle
    '����������㣬�޷����н���
    cur��� = �������_��ɽ(Get����ID(CStr(strҽ����), CStr(gintInsure)))
    If cur����֧�� > cur��� Then
        MsgBox "��Ҫ�ķ����Ѿ�����ʣ�����", vbInformation, gstrSysName
        �������_��ɽ = False
        Exit Function
    End If
    If ������ϸ����(1, lng����ID) = False Then
        �������_��ɽ = False
        Exit Function
    End If
    '���н���׼��
    strSql = "Update Check_doex_interface set Ps_account_pay = " & _
            CStr(cur����֧��) & ",Bala_op_id = '" & ToVarchar(UserInfo.����, 8) & _
            "' where Bill_no = '" & mstr����� & "' and " & _
            "App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    gcn��ɽ.Execute strSql
    
    '�ύ��������
    strSql = "update Check_bill_request set Request_status = '1',Request_Result=null where" & _
            " Bill_no ='" & mstr����� & "' and " & _
            " App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    gcn��ɽ.Execute strSql
    
    'Modified By ���� ���� 06:10:58
    If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
        '����д����������¿�ʧ�ܣ�������������з��ش�������һ���ͻ��������
        Call Shell("D:\hw_ic_write\hw_ic_write.exe " & mstr�����, vbHide)
    End If
    
    If Checkrequest(mstr�����) = False Then �������_��ɽ = False: Exit Function
    
    '���������
    curDate = zlDatabase.Currentdate
    '��ȡ�����ʻ�֧���͸����ֽ�֧��
    strSql = "select Ps_account_pay,Ps_cost_pay,Ps_bala,Plan_pay,acc_cyc from Check_doex_interface" & _
            " where Bill_no ='" & mstr����� & "' and " & _
            " App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
    cur����֧�� = NVL(rs��ɽ("Ps_account_pay"), 0)
    cur��� = NVL(rs��ɽ("Ps_bala"), 0)
    curȫ�Ը� = NVL(rs��ɽ("Ps_cost_pay"), 0)
    str�������� = NVL(rs��ɽ("acc_cyc"), "")
    
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
        cur�ز�ͳ�� = NVL(rs��ɽ("Plan_pay"), 0)
    Else
        cur�ز�ͳ�� = 0
    End If
    curҽ������ = cur�ز�ͳ��
    cur�������� = curȫ�Ը� + cur����֧�� + cur�ز�ͳ��
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(Get����ID(CStr(strҽ����), CStr(gintInsure)), Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & Get����ID(CStr(strҽ����), CStr(gintInsure)) & _
            "," & gintInsure & "," & Year(curDate) & "," & cur�ʻ������ۼ� & _
            "," & cur�ʻ�֧���ۼ� + cur����֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & "," & _
            cur���� & "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & ")"
    Call ExecuteProcedure("��ɽҽ��")
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & _
            Get����ID(CStr(strҽ����), CStr(gintInsure)) & "," & Year(curDate) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� + cur����֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� + cur�ز�ͳ�� & "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & _
            cur�������� & "," & curȫ�Ը� & "," & cur���Ը� & ",NULL," & cur�ز�ͳ�� & ",NULL,NULL," & _
            cur����֧�� & ",NULL)"
    Call ExecuteProcedure("��ɽҽ��")
    
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
'        gstrSQL = "zl_�������ڼ�¼_insert("
        gstrSQL = "Insert into zlhis.�������ڼ�¼ values (" & lng����ID & ",'" & str�������� & "'," & cur�������� & "," & cur����֧�� & "," & cur�ز�ͳ�� & ",'L',to_date('" & Format(curDate, "yyyy-MM-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'))"
        gcnOracle.Execute gstrSQL
'        Call ExecuteProcedure("������ҽ��")
    End If

    strSql = "delete from Check_bill_request  where" & _
            " Bill_no ='" & mstr����� & "' and  App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    gcn��ɽ.Execute strSql
    �������_��ɽ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    �������_��ɽ = False
End Function

Public Function ����������_��ɽ(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, strInput As String, arrOutput  As Variant
    Dim lng����ID As Long, str��ˮ�� As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, str�������� As String
    Dim curƱ���ܽ�� As Currency
    Dim curDate As Date
    
    On Error GoTo errHandle
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ��  From ���˷��ü�¼ Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ɽҽ��")
    
    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ɽҽ��")
    
    lng����ID = rsTemp("����ID")
    
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=" & gintInsure & " and ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ɽҽ��")
    
    If rsTemp.EOF = True Then
        MsgBox "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
'    str��ˮ�� = rsTemp("֧��˳���")
    
'    strInput = "99|" & str��ˮ�� & "|" & ToVarchar(UserInfo.����, 20)
'    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - NVL(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - NVL(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure("��ɽҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        NVL(rsTemp("����ͳ����"), 0) * -1 & "," & NVL(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & NVL(rsTemp("�����Ը����"), 0) & "," & _
        cur�����ʻ� * -1 & ",Null)"
    Call ExecuteProcedure("��ɽҽ��")
    
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
        gstrSQL = "Select * from �������ڼ�¼ where ����id=" & lng����ID
        Call OpenRecordset(rsTemp, "�������")
        If Not rsTemp.EOF Then
            str�������� = rsTemp!��������
    '        gstrSQL = "zl_�������ڼ�¼_insert("
            gstrSQL = "Insert into zlhis.�������ڼ�¼ values (" & lng����ID & ",'" & str�������� & "'," & curƱ���ܽ�� * -1 & "," & cur�����ʻ� * -1 & "," & NVL(rsTemp("ͳ��"), 0) * -1 & ",'L',to_date('" & Format(curDate, "yyyy-MM-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'))"
            gcnOracle.Execute gstrSQL
        End If
'        Call ExecuteProcedure("������ҽ��")
    End If

    ����������_��ɽ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����������_��ɽ(rs������ϸ As Recordset, str���㷽ʽ As String) As Boolean
    Dim cur����֧�� As Currency, cur�����ֽ�֧�� As Currency, cur�����ʻ�֧�� As Currency
    Dim curͳ��֧�� As Currency, cur���֧�� As Currency, lngCount As Long
    Dim strSql As String, rs��ɽ As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset, strBillNO As String
    Dim strMedi As String, strPageId As String
    Dim lng����ID As Long, cur�����ܶ� As Currency
    Dim i As Integer, frm�ȴ� As New frm�ȴ���Ӧ��ɽ
    Dim datCurr As Date, cur�����ʻ���� As Currency
    If Val(Get���ղ���_��ɽ("���õ���")) <> 2 Then          '�������������,�����������
        �����������_��ɽ = False
        Exit Function
    End If
    '�ж��Ƿ��Ѿ���������
    If rs������ϸ.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If
    On Error GoTo errHandle
    '������˵Ĳ�����ҳ��Ҳͬʱ��������㵥��
    lng����ID = rs������ϸ(0)
    strBillNO = mstr�����
'    rs������ϸ.Sort = "�Ƿ��ϴ� desc"
'    ����ϴθ�����ŵ�Ԥ�����¼�������ٴ�����ʱ��������Ų�ͬ������ɾ�������ݱض��Ǳ���δ�����Ԥ������
'    strSql = "delete from Check_item_list_interface where Bill_no = '" & _
'            mstr����� & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
'    gcn��ɽ.Execute strSql
'    strSql = "delete from Check_item_request where Bill_no = '" & _
'            mstr����� & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
'    gcn��ɽ.Execute strSql
    
    '�����ǰ��Ҫ�����
    strSql = "select max(Charge_item_no) as charge_item_no from " & _
            "Check_item_list_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
    If rs��ɽ.EOF Then
        i = 1
    Else
        i = NVL(rs��ɽ("Charge_item_no"), 0) + 1
    End If
    rs������ϸ.MoveFirst
    lngCount = 0
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then Call ShowWindow(frm�ȴ�.hwnd, 9)
    SetPos frm�ȴ�.hwnd
    frm�ȴ�.Move (Screen.Width - frm�ȴ�.Width) / 2, (Screen.Height - frm�ȴ�.Height) / 2
    DoEvents
    Do While Not rs������ϸ.EOF
        '������еķ��ý��
        cur�����ʻ�֧�� = cur�����ʻ�֧�� + rs������ϸ("ʵ�ս��")
        gstrSQL = "Select * From �շ�ϸĿ where id=" & rs������ϸ("�շ�ϸĿID")
        Call OpenRecordset(rsTmp, "������ҽԺ")
        If rsTmp!��� = 5 Or rsTmp!��� = 6 Or rsTmp!��� = 7 Then
            strMedi = "1"
        Else
            strMedi = "2"
        End If
        
        '���������ύ׼��
        strSql = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code," & _
                "App_item_name,Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id)" & _
                " values('" & strBillNO & "','" & _
                Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','" & _
                rs������ϸ("����ID") & "','" & rs������ϸ("������") & _
                "',to_Date('" & Format(Date, "yyyy-MM-dd HH:MM:SS") & _
                "','yyyy-MM-dd HH24:MI:SS'),'" & rs������ϸ("�շ�ϸĿID") & _
                "','Ԥ����','" & strMedi & "','" & _
                rs������ϸ("���㵥λ") & "'," & rs������ϸ("����") & "," & _
                CStr(rs������ϸ("����")) & "," & CStr(rs������ϸ("ʵ�ս��")) & _
                ",to_date('" & Format(Date, "yyyy-MM-dd HH:MM:SS") & _
                "','yyyy-MM-dd HH24:MI:SS'),'" & UserInfo.���� & "')"
        gcn��ɽ.Execute strSql
        
        '�����ύ����
        strSql = "Insert into Check_Item_Request(Bill_no,App_code,Charge_item_no,Request_status) values('" & _
        strBillNO & "','" & Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','0')"
        gcn��ɽ.Execute strSql
        lngCount = lngCount + 1
        '�����ѯ����(�����ڴ�������в��ȴ�����״̬)
'        If frm�ȴ�.Result(2, strBillNo, i) = False Then
'            �����������_��ɽ = False
'            MsgBox "�ڽ���Ĺ���֮�з����ж�", vbInformation, gstrSysName
'            GoTo ResetTrans
'        End If
'        '��ѯ�ύ���
'        strSql = "select Request_Result,Err_Code,Err_text from " & _
'                "check_item_request where Bill_no = '" & strBillNo & _
'                 "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & _
'                 "' and Charge_item_no = '" & CStr(i) & "'"
'        If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
'        rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
'        If rs��ɽ.BOF Then
'            �����������_��ɽ = False
'            GoTo ResetTrans
'        Else
'            If rs��ɽ("Request_Result") = "0" Then
'                MsgBox "��������[" & rs��ɽ("Err_Code") & "]:" & vbCrLf & String(2, "��") & rs��ɽ("Err_text"), vbInformation, gstrSysName
'                �����������_��ɽ = False
'                GoTo ResetTrans
'            End If
'        End If

        '��HIS֮�еĻ������ݽ����޸�
        i = i + 1
        rs������ϸ.MoveNext
    Loop
    Do While True
        '��ѯ�ύ���
        strSql = "select Request_Result,Err_Code,Err_text from " & _
                "check_item_request where Bill_no = '" & strBillNO & _
                 "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & _
                 "' and Request_result is Null"
        If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
        rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
        If rs��ɽ.EOF Then Exit Do
        DoEvents
    Loop
    Unload frm�ȴ�
    cur�����ܶ� = cur�����ʻ�֧��
    '���н���׼��
    strSql = "Update Check_doex_interface set Ps_account_pay = " & _
            CStr(cur����֧��) & ",Bala_op_id = '" & ToVarchar(UserInfo.����, 8) & _
            "' where Bill_no = '" & mstr����� & "' and " & _
            "App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    gcn��ɽ.Execute strSql
    
    '�ύ��������
    strSql = "update Check_bill_request set Request_status = '5',Request_Result=null where" & _
            " Bill_no ='" & strBillNO & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    gcn��ɽ.Execute strSql
    
    If Checkrequest(strBillNO) = False Then
        �����������_��ɽ = False
        GoTo ResetTrans
    End If
    
    '�ӶԷ������ݿ�֮����ȡ�����ʻ�֧�����ֽ�֧����ͳ��֧�������֧��
    strSql = "select Ps_bala from" & _
            " Check_doex_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
    cur�����ʻ�֧�� = NVL(rs��ɽ("Ps_bala"), 0)
    
    strSql = "select Ps_account_pay,Ps_cost_pay,Plan_pay,Big_pay from" & _
            " Check_doex_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
    cur�����ʻ�֧�� = NVL(rs��ɽ("Ps_account_pay"), 0)
    cur�����ֽ�֧�� = NVL(rs��ɽ("Ps_cost_pay"), 0)
    curͳ��֧�� = NVL(rs��ɽ("Plan_pay"), 0)
    cur���֧�� = NVL(rs��ɽ("Big_pay"), 0)
    
'    '������������ʻ�֧��
'    cur�����ܶ� = cur�����ܶ� - curͳ��֧�� - cur���֧��
'    cur�����ʻ�֧�� = IIf(cur�����ʻ�֧�� > cur�����ܶ�, cur�����ܶ�, cur�����ʻ�֧��)
    
    str���㷽ʽ = "�����ʻ�;" & cur�����ʻ�֧�� & ";0" '�����޸ĸ����ʻ�
    If curͳ��֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ = "", "", "|") & "ͳ��֧��;" & curͳ��֧�� & ";0" '�������޸�ͳ��֧��
    End If
    If cur���֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ = "", "", "|") & "���֧��;" & cur���֧�� & ";0" '�������޸Ĵ��֧��
    End If
    �����������_��ɽ = True
ResetTrans:             '�Ժ��ֵ��ݳ��ΪԤ������ϴ��ķ�����ϸ
    '�����ǰ��Ҫ�����
    rs������ϸ.MoveFirst
    strSql = "select max(Charge_item_no) as charge_item_no from " & _
            "Check_item_list_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
    i = NVL(rs��ɽ("Charge_item_no"), 0) + 1
    rs������ϸ.MoveFirst
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then Call ShowWindow(frm�ȴ�.hwnd, 9)
    SetPos frm�ȴ�.hwnd
    frm�ȴ�.Move (Screen.Width - frm�ȴ�.Width) / 2, (Screen.Height - frm�ȴ�.Height) / 2
    DoEvents
    Do While Not rs������ϸ.EOF And lngCount > 0
        '������еķ��ý��
        cur�����ʻ�֧�� = cur�����ʻ�֧�� + rs������ϸ("ʵ�ս��")
        gstrSQL = "Select * From �շ�ϸĿ where id=" & rs������ϸ("�շ�ϸĿID")
        Call OpenRecordset(rsTmp, "������ҽԺ")
        If rsTmp!��� = 5 Or rsTmp!��� = 6 Or rsTmp!��� = 7 Then
            strMedi = "1"
        Else
            strMedi = "2"
        End If
        '���������ύ׼��
        strSql = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code," & _
                "App_item_name,Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id)" & _
                " values('" & strBillNO & "','" & _
                Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','" & _
                rs������ϸ("����ID") & "','" & rs������ϸ("������") & _
                "',to_Date('" & Format(Date, "yyyy-MM-dd HH:MM:SS") & _
                "','yyyy-MM-dd HH24:MI:SS'),'" & rs������ϸ("�շ�ϸĿID") & _
                "','Ԥ����','" & strMedi & "','" & _
                rs������ϸ("���㵥λ") & "'," & 0 - rs������ϸ("����") & "," & _
                CStr(rs������ϸ("����")) & "," & CStr(0 - rs������ϸ("ʵ�ս��")) & _
                ",to_date('" & Format(Date, "yyyy-MM-dd HH:MM:SS") & _
                "','yyyy-MM-dd HH24:MI:SS'),'" & UserInfo.���� & "')"
        gcn��ɽ.Execute strSql
        
        '�����ύ����
        strSql = "Insert into Check_Item_Request(Bill_no,App_code,Charge_item_no,Request_status) values('" & _
        strBillNO & "','" & Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','0')"
        gcn��ɽ.Execute strSql
        lngCount = lngCount - 1
        '�����ѯ����
'        If frm�ȴ�.Result(2, strBillNo, i) = False Then
'            �����������_��ɽ = False
'            MsgBox "�ڽ���Ĺ���֮�з����ж�", vbInformation, gstrSysName
'            Exit Function
'        End If
        '��ѯ�ύ���
        
        i = i + 1
        rs������ϸ.MoveNext
    Loop
    Do While True
        '��ѯ�ύ���
        strSql = "select Request_Result,Err_Code,Err_text from " & _
                "check_item_request where Bill_no = '" & strBillNO & _
                 "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & _
                 "' and Request_result is Null"
        If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
        rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
        If rs��ɽ.EOF Then Exit Do
        DoEvents
    Loop
    
    Unload frm�ȴ�
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    
    �����������_��ɽ = False
End Function

Public Function ������ϸ����(lng��� As Long, Optional lng����ID As Long, Optional strNO As String, Optional lng����ID As Long, Optional int���� As Integer, Optional int״̬ As Integer) As Boolean
'���ܣ�����ύ���������ϸ
'lng��� 1������  2��סԺ
'lng����ID�����������������
'strNo:���ݺ�
'int���ʣ�
'lng����ID  Ĭ��Ϊ0����ʾ�������ŵ��ݣ�����Ϊ������ָ�����˵ġ�����Ҫ����Ϊҽ���ڱ�����ʵ�ʱ���Ƿֲ������ύ���ݶ�����һ���ύ��
    Dim strSql As String, rsTmp As New ADODB.Recordset
    Dim rs��ɽ As New ADODB.Recordset, strBillNO As String
    Dim strMedi As String, i As Integer, rsTemp As New ADODB.Recordset
    Dim frm�ȴ� As New frm�ȴ���Ӧ��ɽ
     
    On Error GoTo errHandle
    If lng����ID = 0 Then
        If lng��� = 1 Then
            gstrSQL = "select ����ID from ���˷��ü�¼ where ����ID = " & _
                    lng����ID & " and rownum < 2"
        Else
            gstrSQL = "select ����ID from ���˷��ü�¼ where NO ='" & _
                    strNO & "' " & " and ��¼���� = " & int���� & _
                    " and ��¼״̬  =" & int״̬ & " and rownum < 2"
        End If
        Call OpenRecordset(rsTmp, "��ɽҽ��")
        lng����ID = rsTmp("����ID")
    End If
    If lng��� = 1 Then
       strBillNO = mstr�����
    Else
        gstrSQL = "select max(��ҳID) as ��ҳID from ������ҳ where ����ID =" & lng����ID
        Call OpenRecordset(rsTmp, "��ɽҽ��")
        strBillNO = CStr(lng����ID) & "_" & CStr(rsTmp("��ҳID"))
    End If
    If lng��� = 1 Then
        '����ǰ���ݵļ�¼�ͼ���¼����ɾ��:ע�⣬�շ�ϸĿ��ν��д��ݻ���Ҫ�޸�
        strSql = "delete from Check_item_list_interface where Bill_no = '" & _
                mstr����� & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSql
        strSql = "delete from Check_item_request where Bill_no = '" & _
                mstr����� & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSql
        gstrSQL = "select A.ID,A.����ʱ��,A.���,A.NO,A.������,A.�Ǽ�ʱ��," & _
                "A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��¼����,A.��¼״̬,D.��Ŀ���� as ϸĿ����,B.���� as ϸĿ����," & _
                "C.����  as ��Ŀ����,B.���㵥λ, (A.���� * A.����) as ����," & _
                "A.��׼����,A.ʵ�ս��,A.����Ա����,A.�Ƿ��ϴ� from  " & _
                "���˷��ü�¼ A,�շ�ϸĿ B,������Ŀ C,����֧����Ŀ D" & _
                " where A.�շ�ϸĿID = B.ID and A.������ĿID = C.ID and A.����ID =" & _
                CStr(lng����ID) & " and A.�շ�ϸĿID = D.�շ�ϸĿID and D.���� = " & _
                gintInsure & " and A.����ID = " & lng����ID
    Else
        gstrSQL = "select A.ID,A.����ʱ��,A.���,A.NO,A.������,A.�Ǽ�ʱ��," & _
                "A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��¼����,A.��¼״̬,D.��Ŀ���� as ϸĿ����,B.���� as ϸĿ����,C.���� as " & _
                "��Ŀ����,B.���㵥λ, (A.���� * A.����) as ����,A.��׼����,A.ʵ�ս��," & _
                "A.����Ա����,A.�Ƿ��ϴ� from ���˷��ü�¼ A,�շ�ϸĿ B,������Ŀ C," & _
                "����֧����Ŀ D where A.�շ�ϸĿID = B.ID and A.������ĿID = C.ID " & _
                " and A.NO ='" & CStr(strNO) & "' and A.��¼״̬ = " & int״̬ & _
                " and A.��¼���� = " & int���� & " and A.�շ�ϸĿID = D.�շ�ϸĿID " & _
                " and D.���� = " & gintInsure & _
                " and A.����ID = " & lng����ID
    End If
    Call OpenRecordset(rsTmp, "��ɽҽ��")
    If rsTmp.BOF Then ������ϸ���� = False: Exit Function
    '�����ʼ���ݵĺ���
    strSql = "select max(Charge_item_no) as Charge_item_no from " & _
            "Check_item_list_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
    If rs��ɽ.EOF Then
        i = 1
    Else
        i = NVL(rs��ɽ("Charge_item_no"), 0) + 1
    End If
    '�𲽽��з�����ϸ����
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then Call ShowWindow(frm�ȴ�.hwnd, 5)
    SetPos frm�ȴ�.hwnd
    frm�ȴ�.Move (Screen.Width - frm�ȴ�.Width) / 2, (Screen.Height - frm�ȴ�.Height) / 2
    DoEvents
    Do While Not rsTmp.EOF
        '���ύ���ݵ�׼��,���Ϊ���ﲡ�˾ʹ��ݡ�����ID + ʱ�䡱�����ΪסԺ���ˣ��ʹ��ݲ���ID����ҳID
        If rsTmp("�վݷ�Ŀ") = "��ҩ��" Or rsTmp("�վݷ�Ŀ") = "�в�ҩ" Or rsTmp("�վݷ�Ŀ") = "�г�ҩ" Then
            strMedi = "1"
        Else
            strMedi = "2"
        End If
        If Val(Get���ղ���_��ɽ("���õ���")) <> 1 Then
            strSql = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                    "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code,App_item_name," & _
                    "Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id)" & _
                    " values('" & strBillNO & "','" & _
                    Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','" & rsTmp("NO") & "','" & _
                    rsTmp("������") & "',to_date('" & Format(rsTmp("�Ǽ�ʱ��"), "yyyy-MM-dd HH:MM:SS") & "','yyyy-MM-dd HH24:MI:SS'),'" & _
                    rsTmp("ϸĿ����") & "','" & rsTmp("ϸĿ����") & "','" & strMedi & _
                    "','" & rsTmp("���㵥λ") & "'," & rsTmp("����") & "," & CStr(rsTmp("��׼����")) & "," & _
                    CStr(rsTmp("ʵ�ս��")) & ",to_date('" & Format(rsTmp("�Ǽ�ʱ��"), "yyyy-MM-dd HH:MM:SS") & "','yyyy-MM-dd HH24:MI:SS'),'" & _
                    rsTmp("����Ա����") & "')"
        Else
            strSql = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                    "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code,App_item_name," & _
                    "Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id)" & _
                    " values('" & strBillNO & "','" & _
                    Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','" & rsTmp("NO") & "','" & _
                    rsTmp("������") & "','" & rsTmp("�Ǽ�ʱ��") & "','" & _
                    rsTmp("ϸĿ����") & "','" & rsTmp("ϸĿ����") & "','" & strMedi & _
                    "','" & rsTmp("���㵥λ") & "'," & rsTmp("����") & "," & CStr(rsTmp("��׼����")) & "," & _
                    CStr(rsTmp("ʵ�ս��")) & ",'" & rsTmp("�Ǽ�ʱ��") & "','" & _
                    rsTmp("����Ա����") & "')"
        End If
        gcn��ɽ.Execute strSql
        '�����ύ����
        strSql = "Insert into Check_item_request(Bill_no,App_code,Charge_item_no,Request_status) values('" & _
                strBillNO & "','" & Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','0')"
        gcn��ɽ.Execute strSql
        '��ѯ�ύ���
        If Val(Get���ղ���_��ɽ("���õ���")) <> 2 Then
            If frm�ȴ�.Result(2, strBillNO, i) = False Then
                ������ϸ���� = False
                MsgBox "������ϸ���ݷ����ж�", vbInformation, gstrSysName
                Exit Function
            End If
            strSql = "select Request_Result,Err_Code,Err_text from check_item_request" & _
                    " where Bill_no = '" & strBillNO & _
                     "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "' and Charge_item_no = '" & _
                     CStr(i) & "'"
            If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
            rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
            If rs��ɽ.BOF Then
                ������ϸ���� = False
                Exit Function
            Else
                If rs��ɽ("Request_Result") = "0" Then
                    MsgBox "��������" & rs��ɽ("Err_Code") & ":" & vbCrLf & String(2, "��") & rs��ɽ("Err_text"), vbInformation, gstrSysName
                    ������ϸ���� = False
                    Exit Function
                End If
            End If
        End If
        '��HIS֮�еĻ������ݽ����޸�
        gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rsTmp("ID") & "')"
        Call ExecuteProcedure("��ɽҽ��")
        rsTmp.MoveNext
        i = i + 1
    Loop
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
        Do While True
            '��ѯ�ύ���
            strSql = "select Request_Result,Err_Code,Err_text from " & _
                    "check_item_request where Bill_no = '" & strBillNO & _
                     "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & _
                     "' and Request_result is Null"
            If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
            rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
            If rs��ɽ.EOF Then Exit Do
            DoEvents
        Loop
        Unload frm�ȴ�
    End If
    rs��ɽ.Close
    ������ϸ���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    ������ϸ���� = False
End Function

Private Function Get����ID(strҽ���� As String, strҽ�����ı��� As String) As String
'���ܣ�ͨ��ҽ�����ĺ����ҽ�����������ID
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select ����ID from �����ʻ� where ���� = '" & strҽ�����ı��� & _
            "' and ҽ���� = '" & strҽ���� & "'"
    Call OpenRecordset(rsTmp, "��ɽҽ��")
    If Not rsTmp.BOF Then
        Get����ID = CStr(rsTmp("����ID"))
    Else
        Get����ID = ""
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Get����ID = ""
End Function

Public Function �������_��ɽ(str����ID As String) As Currency
'���ܣ�ͨ�����˵���Ϣ����������
    Dim strSql As String, rsTmp As New ADODB.Recordset
    Dim strTime As String, rs��ɽ As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    'Modified By ���� ���� 06:06:13
    If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
        '�����������ǭ��������ֱ�Ӵӱ����ʻ��ж�ȡ
        gstrSQL = "Select �ʻ���� ��� From �����ʻ� Where ����ID=" & Val(str����ID)
        Call OpenRecordset(rsTmp, "��ȡ�ʻ����")
        �������_��ɽ = NVL(rsTmp!���, 0)
    Else
        '���������㲻ͨ����ֱ�ӷ���
        gstrSQL = "select ����,���� from �����ʻ� where ����ID = " & str����ID
        Call OpenRecordset(rsTmp, "��ɽҽ��")
        If rsTmp.BOF Then �������_��ɽ = 0: Exit Function
        '�����ݿ�֮�л�ȡ�ֿ����˵���֤��Ϣ
        strTime = CStr(Format(zlDatabase.Currentdate, "yyyymmddhhmmss")) & "00"
        strSql = "insert into Check_doex_interface(Bill_no,App_code," & _
                "Ic_id,Doct_flag,Is_bala,Regi_op_id) values('" & strTime & "','" & Mid(gstrҽԺ����, 1, 4) & "','" & _
                rsTmp("����") & rsTmp("����") & "','0','0','" & ToVarchar(UserInfo.����, 8) & "')"
        gcn��ɽ.Execute strSql
        strSql = "insert into Check_bill_request(Bill_no,App_code," & _
                "Request_status) values('" & strTime & "','" & Mid(gstrҽԺ����, 1, 4) & _
                "','2')"
        gcn��ɽ.Execute strSql
        If Checkrequest(strTime) = False Then �������_��ɽ = 0: Exit Function
        '����Ϣ֮����ȡ���˵ĸ����ʻ����
        strSql = "select Ps_Bala from Check_Doex_Interface where Bill_no = '" & strTime & "'" & _
                " and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
        rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
        If Not rs��ɽ.BOF Then
            �������_��ɽ = IIf(IsNull(rs��ɽ("Ps_Bala")), 0, rs��ɽ("Ps_Bala"))
        Else
            �������_��ɽ = 0
        End If
        strSql = "delete from Check_bill_request where Bill_no = '" & _
                strTime & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSql
        strSql = "delete from Check_doex_interface where Bill_no = '" & _
                strTime & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSql
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    �������_��ɽ = 0
End Function

Public Function ��Ժ�Ǽ�_��ɽ(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim strSql As String, strInNote As String
    Dim rsTmp As New ADODB.Recordset
    
    '������˵������Ϣ
    On Error GoTo errHandle
    gstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,A.��Ժ����,A.סԺҽʦ,C.����," & _
            "C.���� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = " & lng��ҳID & " And A.����ID = " & lng����ID
    Call OpenRecordset(rsTmp, "��ɽҽ��")
    
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID)   '��Ժ���
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 And gstr���ⲡ�� <> "" Then
        '����Ƿ����ⲡ
        strInNote = gstr���ⲡ��
    End If
    If rsTmp.BOF Then ��Ժ�Ǽ�_��ɽ = False: Exit Function
    '׼�������ύ
    strSql = "Delete from Check_doex_interface where bill_no='" & lng����ID & "_" & lng��ҳID & "' and App_code='" & Mid(gstrҽԺ����, 1, 4) & "' and Doct_flag=1 and Hosp_No is null"
    gcn��ɽ.Execute strSql
    strSql = "Delete from Check_bill_request where bill_no='" & lng����ID & "_" & lng��ҳID & "' and App_code='" & Mid(gstrҽԺ����, 1, 4) & "'"
    gcn��ɽ.Execute strSql
    
    If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
        strSql = "Insert into Check_doex_interface(Bill_no,App_code,Doct_flag," & _
                "Doex_no,In_mode,Ill_type,Ic_id,Is_bala,Regi_op_id,Sec_off,The_bunk," & _
                "In_time,Tre_dr) values('" & lng����ID & "_" & lng��ҳID & _
                "','" & Mid(gstrҽԺ����, 1, 4) & "','1','" & NVL(rsTmp("סԺ��")) & "','1','" & _
                strInNote & "','" & NVL(rsTmp("����")) & NVL(rsTmp("����")) & "','0','" & ToVarchar(UserInfo.����, 8) & _
                "','" & NVL(rsTmp("סԺ����")) & "','" & NVL(rsTmp("��Ժ����"), "") & "'," & _
                " '" & NVL(rsTmp("��Ժ����")) & "'" & _
                ",'" & NVL(rsTmp("סԺҽʦ"), "") & "')"
    Else
        strSql = "Insert into Check_doex_interface(Bill_no,App_code,Doct_flag," & _
                "Doex_no,In_mode,Ill_type,Ic_id,Is_bala,Regi_op_id,Sec_off,The_bunk," & _
                "In_time,Tre_dr) values('" & lng����ID & "_" & lng��ҳID & _
                "','" & Mid(gstrҽԺ����, 1, 4) & "','1','" & NVL(rsTmp("סԺ��")) & "','1','" & _
                strInNote & "','" & NVL(rsTmp("����")) & NVL(rsTmp("����")) & "','0','" & ToVarchar(UserInfo.����, 8) & _
                "','" & NVL(rsTmp("סԺ����")) & "','" & NVL(rsTmp("��Ժ����"), "") & "'," & _
                " to_date('" & Format(rsTmp("��Ժ����"), "yyyy-MM-dd HH:MM:SS") & "','yyyy-MM-dd HH24:MI:SS')" & _
                ",'" & NVL(rsTmp("סԺҽʦ"), "") & "')"
    End If
    gcn��ɽ.Execute strSql
    '������Ժ����
    strSql = "Insert into Check_bill_request(Bill_no,App_code,Request_status)" & _
            "values('" & lng����ID & "_" & lng��ҳID & "','" & _
            Mid(gstrҽԺ����, 1, 4) & "','0')"
    gcn��ɽ.Execute strSql
    '��ѯ����Ľ��
    If Checkrequest(lng����ID & "_" & lng��ҳID) = False Then
        ��Ժ�Ǽ�_��ɽ = False
        Exit Function
    End If
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure("��ɽҽ��")
    ��Ժ�Ǽ�_��ɽ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_��ɽ = False
End Function

Public Function ���ʴ���_��ɽ(strNO As String, int���� As Integer, int״̬ As Integer, Optional lng����ID As Long) As Boolean
'��סԺ���˵ķ��ô��ݵ�ҽ������������ͬʱ�޸Ĳ��˷�����Ϣ֮�е�����
    If lng����ID = 0 Then
        ���ʴ���_��ɽ = ������ϸ����(2, , strNO, , int����, int״̬)
    Else
        ���ʴ���_��ɽ = ������ϸ����(2, , strNO, lng����ID, int����, int״̬)
    End If
End Function

Public Function סԺ�������_��ɽ(rs������ϸ As Recordset, lng����ID As Long, strҽ���� As String, str���� As String) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    
    Dim cur�����ʻ�֧�� As Currency, cur�����ֽ�֧�� As Currency
    Dim curͳ��֧�� As Currency, cur���֧�� As Currency, cur�����ܶ� As Currency
    Dim strSql As String, rs��ɽ As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset, strBillNO As String
    Dim strMedi As String, strPageId As String
    Dim i As Integer, frm�ȴ� As New frm�ȴ���Ӧ��ɽ
    Dim datCurr As Date, cur�����ʻ���� As Currency
    
    '�ж��Ƿ��Ѿ���������
    If rs������ϸ.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If
    On Error GoTo errHandle
    '������˵Ĳ�����ҳ��Ҳͬʱ��������㵥��
    gstrSQL = "select max(��ҳID) as ��ҳID from ������ҳ where ����ID =" & lng����ID
    Call OpenRecordset(rsTmp, "��ɽҽ��")
    strPageId = CStr(rsTmp("��ҳID"))
    strBillNO = CStr(lng����ID) & "_" & CStr(rsTmp("��ҳID"))
    rs������ϸ.Sort = "�Ƿ��ϴ� desc"
    '�����ǰ��Ҫ�����
    strSql = "select max(Charge_item_no) as charge_item_no from " & _
            "Check_item_list_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
    If rs��ɽ.EOF Then
        i = 1
    Else
        i = NVL(rs��ɽ("Charge_item_no"), 0) + 1
    End If
    rs������ϸ.MoveFirst
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then Call ShowWindow(frm�ȴ�.hwnd, 5)
    SetPos frm�ȴ�.hwnd

    frm�ȴ�.Move (Screen.Width - frm�ȴ�.Width) / 2, (Screen.Height - frm�ȴ�.Height) / 2
    DoEvents
    Do While Not rs������ϸ.EOF
        '������еķ��ý��
        cur�����ʻ�֧�� = cur�����ʻ�֧�� + rs������ϸ("���")
        '������û�û���ϴ����ͽ����ϴ�:ע�⣬�շ�ϸĿ��ν��д��ݻ���Ҫ�޸�
        
        If IIf(IsNull(rs������ϸ("�Ƿ��ϴ�")), "0", rs������ϸ("�Ƿ��ϴ�")) = "0" Then
            gstrSQL = "select A.ID,A.����ʱ��,A.���,A.NO,A.������,A.�Ǽ�ʱ��," & _
                    "A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��¼����,A.��¼״̬,D.��Ŀ���� as ϸĿ����,B.���� as ϸĿ����,C.����" & _
                    " as ��Ŀ����,B.���㵥λ, (A.���� * A.����) as ����," & _
                    "A.��׼����,A.ʵ�ս��,A.����Ա���� from ���˷��ü�¼ A," & _
                    "�շ�ϸĿ B,������Ŀ C,����֧����Ŀ D where A.�շ�ϸĿID = B.ID and " & _
                    "A.������ĿID = C.ID " & " And A.����ID=" & lng����ID & _
                    " and A.NO ='" & CStr(rs������ϸ("NO")) & "' and " & _
                    "A.��¼״̬ = " & rs������ϸ("��¼״̬") & " and " & _
                    "A.��¼���� = " & rs������ϸ("��¼����") & _
                    " and (A.�۸񸸺� = " & rs������ϸ("���") & " or A.�۸񸸺� Is Null And A.���=" & rs������ϸ("���") & ")" & _
                    " and (A.�Ƿ��ϴ� = 0 or A.�Ƿ��ϴ� is null) and " & _
                    "A.�շ�ϸĿID = D.�շ�ϸĿID and D.���� = " & gintInsure
            Call OpenRecordset(rsTmp, "��ɽҽ��")

            If Not rsTmp.BOF Then
                If rsTmp("�վݷ�Ŀ") = "��ҩ��" Or rsTmp("�վݷ�Ŀ") = "�в�ҩ" Or rsTmp("�վݷ�Ŀ") = "�г�ҩ" Then
                    strMedi = "1"
                Else
                    strMedi = "2"
                End If
                '���������ύ׼��
                If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
                    strSql = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                            "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code," & _
                            "App_item_name,Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id)" & _
                            " values('" & strBillNO & "','" & _
                            Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','" & _
                            rsTmp("NO") & "','" & rsTmp("������") & _
                            "','" & rsTmp("�Ǽ�ʱ��") & _
                            "','" & rsTmp("ϸĿ����") & _
                            "','" & rsTmp("ϸĿ����") & "','" & strMedi & "','" & _
                            rsTmp("���㵥λ") & "'," & rsTmp("����") & "," & _
                            CStr(rsTmp("��׼����")) & "," & CStr(rsTmp("ʵ�ս��")) & _
                            ",'" & rsTmp("�Ǽ�ʱ��") & _
                            "','" & rsTmp("����Ա����") & "')"
                Else
                    strSql = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                            "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code," & _
                            "App_item_name,Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id)" & _
                            " values('" & strBillNO & "','" & _
                            Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','" & _
                            rsTmp("NO") & "','" & rsTmp("������") & _
                            "',to_Date('" & Format(rsTmp("�Ǽ�ʱ��"), "yyyy-MM-dd HH:MM:SS") & _
                            "','yyyy-MM-dd HH24:MI:SS'),'" & rsTmp("ϸĿ����") & _
                            "','" & rsTmp("ϸĿ����") & "','" & strMedi & "','" & _
                            rsTmp("���㵥λ") & "'," & rsTmp("����") & "," & _
                            CStr(rsTmp("��׼����")) & "," & CStr(rsTmp("ʵ�ս��")) & _
                            ",to_date('" & Format(rsTmp("�Ǽ�ʱ��"), "yyyy-MM-dd HH:MM:SS") & _
                            "','yyyy-MM-dd HH24:MI:SS'),'" & rsTmp("����Ա����") & "')"
                End If
                gcn��ɽ.Execute strSql
                '�����ύ����
                strSql = "Insert into Check_Item_Request(Bill_no,App_code,Charge_item_no,Request_status) values('" & _
                strBillNO & "','" & Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','0')"
                gcn��ɽ.Execute strSql
                '�����ѯ����
                If Val(Get���ղ���_��ɽ("���õ���")) <> 2 Then
                    If frm�ȴ�.Result(2, strBillNO, i) = False Then
                        סԺ�������_��ɽ = ""
                        MsgBox "�ڽ���Ĺ���֮�з����ж�", vbInformation, gstrSysName
                        Exit Function
                    End If
                    '��ѯ�ύ���
                    strSql = "select Request_Result,Err_Code,Err_text from " & _
                            "check_item_request where Bill_no = '" & strBillNO & _
                             "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & _
                             "' and Charge_item_no = '" & CStr(i) & "'"
                    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
                    rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
                    If rs��ɽ.BOF Then
                        סԺ�������_��ɽ = ""
                        Exit Function
                    Else
                        If rs��ɽ("Request_Result") = "0" Then
                            MsgBox "��������[" & rs��ɽ("Err_Code") & "]:" & vbCrLf & String(2, "��") & rs��ɽ("Err_text"), vbInformation, gstrSysName
                            סԺ�������_��ɽ = ""
                            Exit Function
                        End If
                    End If
                End If
                '��HIS֮�еĻ������ݽ����޸�
                gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rsTmp("ID") & "')"
                Call ExecuteProcedure("��ɽҽ��")
            End If
            i = i + 1
        End If
        rs������ϸ.MoveNext
    Loop
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
'        Do While True
'            '��ѯ�ύ���
'            strSql = "select Request_Result,Err_Code,Err_text from " & _
'                    "check_item_request where Bill_no = '" & strBillNo & _
'                     "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & _
'                     "' and Request_result is Null"
'            If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
'            rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
'            If rs��ɽ.EOF Then Exit Do
'            DoEvents
'        Loop
        Unload frm�ȴ�
    End If
    cur�����ܶ� = cur�����ʻ�֧��
    If Val(Get���ղ���_��ɽ("���õ���")) <> 1 Then
        '�����ύ׼��
        datCurr = zlDatabase.Currentdate
        strSql = "Update Check_doex_interface set Ps_account_pay = " & _
                cur�����ʻ�֧�� & ",Bala_op_id = '" & ToVarchar(UserInfo.����, 8) & _
                "',Out_time =to_date('" & Format(datCurr, "yyyy-MM-dd") & "','yyyy-MM-dd') " & _
                "where Bill_no = '" & strBillNO & "' and App_code = '" & _
                Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSql
        '���������������,Ŀǰ����֪������Ĳ���ֵ,�ڱ���֮����Ҫ�����޸�
        strSql = "Update Check_bill_request set Request_status = '2',Request_Result=null where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSql
        If Checkrequest(strBillNO) = False Then
            סԺ�������_��ɽ = ""
            Exit Function
        End If
        strSql = "select Ps_bala from" & _
                " Check_doex_interface where Bill_no = '" & strBillNO & _
                "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
        rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
        cur�����ʻ�֧�� = NVL(rs��ɽ("Ps_bala"), 0)
        
        strSql = "Update Check_bill_request set Request_status = '5',Request_Result=null where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSql
        If Checkrequest(strBillNO) = False Then
            סԺ�������_��ɽ = ""
            Exit Function
        End If
    Else
        MsgBox "������ֹ����㣬������ɺ�����ȷ��������......", vbInformation, "ҽҵ���"
    End If
    
    '�ӶԷ������ݿ�֮����ȡ�����ʻ�֧�����ֽ�֧����ͳ��֧�������֧��
    strSql = "select Ps_account_pay,Ps_cost_pay,Plan_pay,Big_pay from" & _
            " Check_doex_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
    If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
        cur�����ʻ�֧�� = NVL(rs��ɽ("Ps_account_pay"), 0)            '��ɽ���ظ����ʻ�֧��
    End If
    cur�����ֽ�֧�� = NVL(rs��ɽ("Ps_cost_pay"), 0)
    curͳ��֧�� = NVL(rs��ɽ("Plan_pay"), 0)
    cur���֧�� = NVL(rs��ɽ("Big_pay"), 0)
    
    '������������ʻ�֧�����
    If Val(Get���ղ���_��ɽ("���õ���")) <> 1 Then
        cur�����ܶ� = cur�����ܶ� - curͳ��֧�� - cur���֧��
        cur�����ʻ�֧�� = IIf(cur�����ʻ�֧�� > cur�����ܶ�, cur�����ܶ�, cur�����ʻ�֧��)
    End If
'    gstrSQL = "Select Nvl(�ʻ����,0) ��� From �����ʻ� Where ����ID=" & lng����ID
'    Call OpenRecordset(rsTmp, "��ȡ�ʻ����")
'    cur�����ʻ���� = rsTmp!���
    
'    If cur�����ʻ�֧�� <> 0 Then
        סԺ�������_��ɽ = "�����ʻ�;" & cur�����ʻ�֧�� & ";0" '�������޸ĸ����ʻ�
'    End If
'    If סԺ�������_��ɽ = "" Then סԺ�������_��ɽ = "�����ʻ�;" & 0 & ";1"
    If curͳ��֧�� <> 0 Then
        סԺ�������_��ɽ = סԺ�������_��ɽ & IIf(סԺ�������_��ɽ = "", "", "|") & "ͳ��֧��;" & curͳ��֧�� & ";0" '�������޸�ͳ��֧��
    End If
    If cur���֧�� <> 0 Then
        סԺ�������_��ɽ = סԺ�������_��ɽ & IIf(סԺ�������_��ɽ = "", "", "|") & "���֧��;" & cur���֧�� & ";0" '�������޸Ĵ��֧��
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Resume
    סԺ�������_��ɽ = ""
End Function

Public Function סԺ����_��ɽ(lng����ID As Long, ByVal lng����ID As Long) As Boolean
'�����˵ķ��ý��н��㣬���ڱ�ɽҽ������Ҫ���г�Ժ�Ǽǣ���˲����г�Ժ�Ǽ�
    Dim rsTmp As New ADODB.Recordset, cur������ As Currency
    Dim strBillNO As String, strSql As String, datCurr As Date
    Dim rs��ɽ As New ADODB.Recordset, cur�����ʻ�֧�� As Currency
    Dim cur�����ֽ�֧�� As Currency, curͳ��֧�� As Currency
    Dim cur���֧�� As Currency, intסԺ�����ۼ� As Integer
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim curͳ���Ը� As Currency, cur�����Ը� As Currency
    Dim cur�����Ը� As Currency, cur��ͳ�� As Currency
    Dim cur���Ը� As Currency, cur���� As Currency
    Dim curȫ�Ը� As Currency, cur�ҹ��Ը� As Currency
    Dim cur�����ʻ� As Currency, str�������� As String
    
    On Error GoTo errHandle
    gstrSQL = "select sum(ʵ�ս��) as ������,sum(���ʽ��) as �ѽ��� from ���˷��ü�¼ where " & _
            "����ID=" & lng����ID & " and ����ID=" & lng����ID
    Call OpenRecordset(rsTmp, "��ɽҽ��")
    cur������ = NVL(rsTmp("�ѽ���"), 0)
    gstrSQL = "select ��ҳID,��Ժ���� from ������ҳ where ��ҳID=(select max(��ҳID) from " & _
            "������ҳ where ����ID  = " & lng����ID & ") and ����ID = " & lng����ID
    Call OpenRecordset(rsTmp, "��ɽҽ��")
    If rsTmp.BOF Then Exit Function
    strBillNO = lng����ID & "_" & rsTmp("��ҳID")
    If Val(Get���ղ���_��ɽ("���õ���")) <> 1 Then
        '�����ύ׼��
        
        strSql = "Update Check_doex_interface set Ps_account_pay = " & cur������ & _
                ",Bala_op_id = '" & ToVarchar(UserInfo.����, 8) & "',Out_time = to_date('" & _
                Format(rsTmp("��Ժ����"), "yyyy-MM-dd") & "','yyyy-MM-dd') where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSql
        '���н�������
        strSql = "Update Check_bill_request set Request_status = '1',Request_Result=null where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSql
        If Checkrequest(strBillNO) = False Then סԺ����_��ɽ = False: Exit Function
    End If
    '�������
    'modify by ccy, add select field Ps_bala
    strSql = "select Ps_bala,Ps_account_pay,Ps_cost_pay,Plan_pay,Big_pay,acc_cyc from" & _
            " Check_doex_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
    cur�����ʻ�֧�� = NVL(rs��ɽ("Ps_account_pay"), 0)
    cur�����ֽ�֧�� = NVL(rs��ɽ("Ps_cost_pay"), 0)
    curͳ��֧�� = NVL(rs��ɽ("Plan_pay"), 0)
    cur���֧�� = NVL(rs��ɽ("Big_pay"), 0)
    cur��ͳ�� = cur���֧��
    curȫ�Ը� = cur�����ʻ�֧��
    cur�����ʻ� = NVL(rs��ɽ("Ps_bala"), 0)
    str�������� = NVL(rs��ɽ("ACC_CYC"), "")
    '��д�����
    datCurr = zlDatabase.Currentdate
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
            
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(datCurr) & "," & _
            cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & _
            cur����ͳ���ۼ� + curͳ��֧�� + curͳ���Ը� + cur�����Ը� + cur�����Ը� + cur��ͳ�� + cur���Ը� & "," & _
            curͳ�ﱨ���ۼ� + curͳ��֧�� + cur��ͳ�� & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure("��ɽҽ��")
    
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & ",NULL," & cur�����Ը� & "," & _
        cur������ & "," & cur�����ֽ�֧�� & "," & cur�ҹ��Ը� & "," & _
        curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & "," & cur���Ը� & "," & cur�����Ը� & "," & cur�����ʻ�֧�� & ",'')"
    Call ExecuteProcedure("��ɽҽ��")
    
    '���ս������
    gstrSQL = "zl_���ս������_insert(" & lng����ID & ",0," & curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & ",NULL)"
    Call ExecuteProcedure("��ɽҽ��")
    
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
'        gstrSQL = "zl_�������ڼ�¼_insert(" & lng����ID & ",'" & str�������� & "'," & cur������ & "," & cur�����ʻ�֧�� & "," & curͳ��֧�� & ",'N',to_date('" & datCurr & "','yyyy-mm-dd HH:MI:SS'))"
        gstrSQL = "Insert into zlhis.�������ڼ�¼ values (" & lng����ID & ",'" & str�������� & "'," & cur������ & "," & cur�����ʻ�֧�� & "," & curͳ��֧�� & ",'N',to_date('" & Format(datCurr, "yyyy-MM-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'))"
        gcnOracle.Execute gstrSQL
'        Call ExecuteProcedure("������ҽ��")
    End If
    
    סԺ����_��ɽ = True
    'modify by ccy
    If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
        MsgBox "���ĸ����ʻ����Ϊ[" & Format(cur�����ʻ�, "0.00") & "Ԫ]", vbInformation, "סԺ����"
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    סԺ����_��ɽ = False
End Function

Public Function סԺ�������_��ɽ(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset, strInput As String, arrOutput  As Variant
    Dim lng����ID As Long, str��ˮ�� As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, str�������� As String
    Dim curDate As Date
        
    On Error GoTo errHandle
    curDate = zlDatabase.Currentdate
    
    '�˷�
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=" & lng����ID
    Call OpenRecordset(rsTemp, "�������")
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    'Ϊ�˽���ʱд���Ľ����������ٴη��ʼ�¼
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=" & gintInsure & " and ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, "�������")
    If rsTemp.EOF = True Then
        MsgBox "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
'    If CanסԺ�������(rsTemp("����ID"), rsTemp("��ҳID")) = False Then Exit Function
    
'    str��ˮ�� = NVL(rsTemp("֧��˳���"), "0")
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(rsTemp("����ID"), Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & rsTemp("����ID") & "," & gintInsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - rsTemp("�����ʻ�֧��") & "," & cur����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure("��ɽҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & rsTemp("����ID") & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - rsTemp("�����ʻ�֧��") & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & rsTemp("�������ý��") * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0,0," & _
        rsTemp("�����ʻ�֧��") * -1 & ",Null," & rsTemp("��ҳID") & "," & rsTemp("��;����") & ")"
    Call ExecuteProcedure("��ɽҽ��")
    
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
        gstrSQL = "Select * from �������ڼ�¼ where ����id=" & lng����ID
        Call OpenRecordset(rsTemp, "�������")
        If Not rsTemp.EOF Then
            str�������� = rsTemp!��������
    '        gstrSQL = "zl_�������ڼ�¼_insert(" & lng����ID & ",'" & str�������� & "'," & NVL(rsTemp("�������ý��"), 0) * -1 & "," & NVL(rsTemp("�����ʻ�֧��"), 0) * -1 & "," & NVL(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",'N',to_date('" & curDate & "','yyyy-mm-dd HH:MI:SS'))"
    '        Call ExecuteProcedure("������ҽ��")
            gstrSQL = "Insert into zlhis.�������ڼ�¼ values (" & lng����ID & ",'" & str�������� & "'," & NVL(rsTemp("�ܶ�"), 0) * -1 & "," & NVL(rsTemp("����"), 0) * -1 & "," & NVL(rsTemp("ͳ��"), 0) * -1 & ",'N',to_date('" & Format(curDate, "yyyy-MM-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'))"
            gcnOracle.Execute gstrSQL
        End If
    End If

    סԺ�������_��ɽ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_��ɽ(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    '����״̬���޸�
    Dim strSql As String, rs��ɽ As New ADODB.Recordset
    Dim strBillNO As String, rsTmp As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim bln����ó�Ժ As Boolean
    
    On Error GoTo errHandle
    '���ô�סԺ�Ƿ�û�з��÷���
    gstrSQL = "Select sum(ʵ�ս��) as ���  from ���˷��ü�¼ where ����ID=" & lng����ID & " and ��ҳID=" & lng��ҳID
    Call OpenRecordset(rsTemp, "���˳�Ժ")
    If rsTemp.EOF = True Then
        bln����ó�Ժ = True
    Else
        bln����ó�Ժ = (NVL(rsTemp("���"), 0) = 0)
    End If
    
    If bln����ó�Ժ = True Then
        '��������ó�Ժ���ͽ��䴦��Ϊ����Ժ�������ø�����סԺ��Ϣ
        gstrSQL = "select ��Ժ���� from ������ҳ where ����ID = " & lng����ID & _
                " and ��ҳID=" & lng��ҳID
        Call OpenRecordset(rsTmp, "��ɽҽ��")
        strBillNO = lng����ID & "_" & lng��ҳID
        '���г�Ժ����
        strSql = "Update Check_bill_request set Request_status= '3',Request_Result=null where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & _
                Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSql
        '��ѯ������
        If Checkrequest(strBillNO) = False Then ��Ժ�Ǽ�_��ɽ = False: Exit Function
        
        'ɾ�����ε���Ժ�Ǽ���Ϣ
        strSql = "Delete from Check_doex_interface where bill_no='" & lng����ID & "_" & lng��ҳID & "' and App_code='" & Mid(gstrҽԺ����, 1, 4) & "' and Doct_flag=1"
        gcn��ɽ.Execute strSql
        strSql = "Delete from Check_bill_request where bill_no='" & lng����ID & "_" & lng��ҳID & "' and App_code='" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSql
    End If
    '��HIS֮�еĻ������ݽ����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure("��ɽҽ��")
    ��Ժ�Ǽ�_��ɽ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    ��Ժ�Ǽ�_��ɽ = False
End Function

Public Function Checkrequest(strBillNO As String) As Boolean
'���ܣ��ж��Ƿ��ܹ������ȷ�Ĳ�����Ϣ
    Dim strSql As String, rs��ɽ As New ADODB.Recordset
    Dim strResult As String '����Ľ��
    Dim strTmp As String, strError As String
    Dim frm�ȴ� As New frm�ȴ���Ӧ��ɽ, lngErrLine As Long
    
    On Error GoTo errHandle
    '�ύ���󣬽��в�ѯ
    If frm�ȴ���Ӧ��ɽ.Result(1, strBillNO) = False Then
        Checkrequest = False: lngErrLine = 1
        Unload frm�ȴ���Ӧ��ɽ
        DoEvents
        Exit Function
    End If
    Unload frm�ȴ���Ӧ��ɽ
    '���ݷ��صķ��ص�ֵ�жϽ��
    strSql = "Select Request_Result,Err_text from " & _
            "Check_bill_request where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'": lngErrLine = 2
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close: lngErrLine = 3
    rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly: lngErrLine = 4
    If Not rs��ɽ.BOF Then
        strTmp = NVL(rs��ɽ("Request_Result"), 0): lngErrLine = 5
        strError = NVL(rs��ɽ("Err_text"), ""): lngErrLine = 6
    Else
        Exit Function
    End If
    Select Case strTmp
        Case "0"
            MsgBox "û�������������������", vbInformation, gstrSysName
            Checkrequest = False
            Exit Function
        Case "1"
            If strError <> "" Then
                MsgBox "ҽ���ӿڵ��ó�����������" & vbCrLf & vbCrLf & strError, vbInformation, gstrSysName
            Else
                MsgBox "ҽ���ӿڵ��ó��ִ���", vbInformation, gstrSysName
            End If
            Exit Function
        Case "9"
            Checkrequest = True
    End Select
    Checkrequest = True
    Exit Function
errHandle:
    MsgBox "�ڹ���[CheckRequest]�е�" & lngErrLine & "�з�������", vbExclamation, "����"
    If ErrCenter() = 1 Then
        Resume
    End If
    Checkrequest = False
End Function

Public Function Get���ղ���_��ɽ(ByVal str������ As String) As String
'���ܣ���ñ��ղ���
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.������,A.����ֵ from ���ղ��� A " & _
              " where A.������='" & str������ & "' and A.����=" & TYPE_�����ɽ & " and A.���� is null "
    Call OpenRecordset(rsTemp, "��ɽҽ��")
    
    If rsTemp.EOF = False Then
        Get���ղ���_��ɽ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
    End If
End Function

Public Sub SetPos(lHwnd As Long, Optional TopFlag As Boolean = True)
    If TopFlag Then
        SetWindowPos lHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    Else
        SetWindowPos lHwnd, HWND_NOTTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    End If
End Sub

