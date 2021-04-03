Attribute VB_Name = "mdl��Ҧ"
Option Explicit
    
Public Declare Function f_Init Lib "dhpDLL.DLL" () As Integer
Public Declare Function f_Close Lib "dhpDLL.DLL" () As Integer
Public Declare Function f_Apply Lib "dhpDLL.DLL" (ByVal lngTradeTypeID As Long, _
    ByVal dblTradeID As Double, ByVal strData As String, ByRef strMessage As String) As Integer

Public gstrOutput��Ҧ As String, gstrInput��Ҧ As String, gcn��Ҧ As New ADODB.Connection, gstrIC���� As String
Private mstrBillNo As String, mcur��ҽ�� As Currency, mstr��ˮ�� As String

Public Function makeBillNO(lng����ID As Long) As String
    Dim datCurr As Date
    datCurr = zlDatabase.Currentdate
    makeBillNO = toHex(CDbl(Format(datCurr, "yyyymmddHHMMSS") & lng����ID), 36)
End Function

Public Function makeICInfo(lng����ID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    '����IC����
    gstrSQL = "Select A.����,B.����,B.�Ա�,A.��λ���� From �����ʻ� A,������Ϣ B Where A.����ID=" & lng����ID & _
        " And A.����=" & gintInsure & " And A.����ID=B.����ID"
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "û���ҵ��ò��˵������Ϣ", vbInformation, gstrSysName
        Exit Function
    End If
    makeICInfo = Right(Space(18) & rsTemp(0), 18) & _
                 String(18, "0") & _
                 Right(Space(20) & rsTemp(1), 20) & _
                 Right(Space(2) & rsTemp(2), 2) & _
                 String(56, "0") & _
                 Right(Space(10) & rsTemp(3), 10) & _
                 String(2, "0") & String(126 + 85 + 146 + 116 * 6, "0")
End Function

Private Function toHex(ByVal dblNum As Double, Optional ByVal dblKey As Double = 16) As String
    Dim dblTemp As Double, dblMod As Double, strTemp As String
    dblTemp = dblNum
    Do
        dblMod = dblTemp - Int(dblTemp / dblKey) * dblKey
        dblTemp = Int(dblTemp / dblKey)
        If dblMod >= 10 Then
            strTemp = Chr(dblMod + 55) & strTemp
        Else
            strTemp = dblMod & strTemp
        End If
    Loop While dblTemp >= dblKey
    dblMod = dblTemp
    If dblMod >= 10 Then
        strTemp = Chr(dblMod + 55) & strTemp
    Else
        strTemp = IIf(dblMod <> 0, dblMod, "") & strTemp
    End If
    toHex = strTemp
End Function

Public Function CheckReturn_��Ҧ() As Boolean
    If glngReturn < 0 Then
        If Split(gstrOutput��Ҧ, "$$")(1) = "" Then
            MsgBox "����ҽ������ʱ��������", vbInformation, gstrSysName
        Else
            MsgBox "ҽ�������������´���" & vbCrLf & "    " & Split(gstrOutput��Ҧ, "$$")(1), vbInformation, gstrSysName
        End If
        Exit Function
    End If
    CheckReturn_��Ҧ = True
End Function

Public Function ���뽻����ˮ_��Ҧ(str�������� As String) As String
    Dim strTemp As String
    ���뽻����ˮ_��Ҧ = ""
    strTemp = "$$" & str�������� & "$$"
    glngReturn = f_Apply(23, 0, strTemp, gstrOutput��Ҧ)
    If CheckReturn_��Ҧ() = False Then Exit Function
    ���뽻����ˮ_��Ҧ = Split(gstrOutput��Ҧ, "$$")(2)
End Function

Public Function openConn��Ҧ() As Boolean
    Dim rsTemp As New ADODB.Recordset, strServer As String, strUser As String, strPass As String, _
        strTemp As String, strDatabase As String
    On Error GoTo errH
    If gcn��Ҧ.State <> adStateOpen Then
        '���ȶ���������������
        gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & gintInsure
        Call OpenRecordset(rsTemp, gstrSysName)
        Do Until rsTemp.EOF
            strTemp = Nvl(rsTemp("����ֵ"), "")
            Select Case rsTemp("������")
                Case "��Ҧ������"
                    strServer = strTemp
                Case "��Ҧ�û���"
                    strUser = strTemp
                Case "��Ҧ�û�����"
                    strPass = strTemp
                Case "��Ҧ���ݿ�"
                    strDatabase = strTemp
            End Select
            rsTemp.MoveNext
        Loop
    
        On Error Resume Next
        gcn��Ҧ.ConnectionString = "Provider=SQLOLEDB.1;Initial Catalog=" & strDatabase & ";Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
        gcn��Ҧ.CursorLocation = adUseClient
        gcn��Ҧ.Open
        If Err.Number <> 0 Then
            MsgBox "ҽ��ǰ�÷���������ʧ�ܡ�", vbInformation, gstrSysName
            openConn��Ҧ = False
            Exit Function
        End If
        On Error GoTo errH
    End If
    openConn��Ҧ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    openConn��Ҧ = False
End Function

Public Function ҽ����ʼ��_��Ҧ() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    If openConn��Ҧ() = False Then
        ҽ����ʼ��_��Ҧ = False
        Exit Function
    End If
    
    gstrInput��Ҧ = "$$$$": gstrOutput��Ҧ = "$$$$$$"
    glngReturn = f_Init()
    ҽ����ʼ��_��Ҧ = CheckReturn_��Ҧ()
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    ҽ����ʼ��_��Ҧ = False
End Function

Public Function ҽ����ֹ_��Ҧ() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    On Error GoTo errHandle
    Set gcn��Ҧ = Nothing
    glngReturn = f_Close()
    ҽ����ֹ_��Ҧ = CheckReturn_��Ҧ()
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ҽ����ֹ_��Ҧ = False
End Function

Public Function ҽ������_��Ҧ() As Boolean
    ҽ������_��Ҧ = frmSet��Ҧ.��������()
End Function

Public Function �����������_��Ҧ(rs������ϸ As Recordset, str���㷽ʽ As String) As Boolean
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
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim strҽ���� As String, lng����ID As Long, datCurr As Date, rsTemp As New ADODB.Recordset, str��Ŀ���� As String, _
        str����ID As String, str���� As String, strSql As String, strTemp As String, iLoop As Long, lng��ˮ As Long, _
        strҽ������ As String, str��ϸ���� As String, str��Ŀ���� As String, str��� As String, str�Ը����� As String
    WriteInfo vbCrLf & "����Ԥ����"
    On Error GoTo errHandle
    If rs������ϸ.RecordCount = 0 Then
        MsgBox "û�в��˷�����ϸ�����ܽ���ҽ������", vbInformation, gstrSysName
        Exit Function
    End If
    rs������ϸ.MoveFirst
    lng����ID = rs������ϸ!����ID
    datCurr = zlDatabase.Currentdate
    
    'ǿ��ѡ����
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & gintInsure

    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ������")
    If rsTemp.State = 1 Then
        str����ID = rsTemp!����
        str���� = rsTemp!����
    Else
        �����������_��Ҧ = False
        Exit Function
    End If
    
    mstrBillNo = makeBillNO(lng����ID)
    gstrSQL = "Select * From �����ʻ� Where ����=" & gintInsure & " And ����ID=" & lng����ID
'    Call OpenRecordset(rsTemp, gstrSysName)
    Set rsTemp = gcnOracle.Execute(gstrSQL)
    
    If rsTemp.EOF Then
        MsgBox "û���ҵ��ò��˵�ҽ����Ϣ", vbInformation, gstrSysName
        Exit Function
    End If
    strҽ���� = rsTemp!����
    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(29)
    If mstr��ˮ�� = "" Then Exit Function
    'д������
    strSql = "Insert Into hi_ClinicRx (BillID,DateDiagnose,ChargeType,HospitalID,PIN,ClinicSerial,Department,DepartmentID," & _
        "Doctor,Disease,DiseaseID,Description,DateOccur,Operator) values ('" & mstr��ˮ�� & "','" & Format(datCurr, "yyyy-mm-dd HH:MM:SS") & _
        "',1," & Trim(gstrҽԺ����) & ",'" & strҽ���� & "','" & lng����ID & "',Null,Null,'" & rs������ϸ!������ & _
        "','" & str���� & "','" & str����ID & "',Null,'" & Format(datCurr, "yyyy-mm-dd HH:MM:SS") & _
        "','" & UserInfo.���� & "')"
    WriteInfo "дǰ�û���������:" & strSql
    gcn��Ҧ.Execute strSql
    mcur��ҽ�� = 0
    iLoop = 1
    strSql = "Select Max(SerialNum) From hi_ClinicPrescription"
    Set rsTemp = gcn��Ҧ.Execute(strSql)
    If rsTemp.EOF Then
        lng��ˮ = 0
    Else
        lng��ˮ = Nvl(rsTemp(0), 0)
    End If
    
    While Not rs������ϸ.EOF
        'ȡ�շ���ϸ
        gstrSQL = "Select ����,����,���,nvl(���,'') as ��� From �շ�ϸĿ Where ID=" & rs������ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, gstrSysName)
        str��ϸ���� = rsTemp!����: str��Ŀ���� = rsTemp!����
        str��� = Left(Left(rsTemp!��� & " |", InStr(rsTemp!��� & " |", "|") - 1), InStr(rsTemp!��� & " |", " ") - 1)
        '�ж���Ŀ����
        str��Ŀ���� = IIf(rsTemp!��� = "5" Or rsTemp!��� = "6", "ҩƷ", IIf(rsTemp!��� = "7", "��ҩ", "����"))
        
        '�ӱ���֧����Ŀ�в����Ƿ��и�ҽ����Ŀ
        gstrSQL = "Select ��Ŀ����,��Ŀ���� From ����֧����Ŀ Where �Ƿ�ҽ��=1 And ����=" & gintInsure & " And �շ�ϸĿID=" & rs������ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, gstrSysName)
        If rsTemp.EOF Then      'û����Ŀ����
            mcur��ҽ�� = mcur��ҽ�� + rs������ϸ!ʵ�ս��
            If str��Ŀ���� = "ҩƷ" Then    '����ΪҩƷʱ��ҽ������Ϊ�����ࡱ
                strҽ������ = "����": str�Ը����� = "1"
            Else        '��Ŀ����Ϊ���ƻ���ҩʱ��ҽ������Ϊ�����ࡱ
                strҽ������ = "����": str�Ը����� = "0"
            End If
        Else            '�и���Ŀʱ����
            str��ϸ���� = rsTemp!��Ŀ����
            If str��Ŀ���� = "����" Then
                gstrSQL = "Select * From hi_Diagnose Where DiagnoseID='" & str��ϸ���� & "'"
            Else
                gstrSQL = "Select * From hi_Medicine Where MedicineID='" & str��ϸ���� & "'"
            End If
            Set rsTemp = gcn��Ҧ.Execute(gstrSQL)
            If rsTemp.EOF Then      '���ҽ������Ŀ¼��δ�ҵ�����Ŀ
                If str��Ŀ���� = "ҩƷ" Then    '����ΪҩƷʱ��ҽ������Ϊ�����ࡱ
                    strҽ������ = "����": str�Ը����� = "1"
                Else        '��Ŀ����Ϊ���ƻ���ҩʱ��ҽ������Ϊ�����ࡱ
                    strҽ������ = "����": str�Ը����� = "0"
                End If
            Else        '���ҽ������Ŀ¼���и�ҩƷ
                strҽ������ = IIf(rsTemp!zfbl = 0, "����", IIf(rsTemp!zfbl = 1, "����", "����"))
                str�Ը����� = rsTemp!zfbl
            End If
        End If
        strSql = "Insert Into hi_ClinicPrescription (SerialNum,HospitalID,BillID,DateDiagnose,RecipeSerial,Class,ItemID,ItemName," & _
            "Specification,Price,Dosage,ScaleSelf,Operator) Values (" & iLoop + lng��ˮ & "," & Trim(gstrҽԺ����) & ",'" & mstrBillNo & _
            "','" & Format(datCurr, "yyyy-mm-dd HH:MM:SS") & "','" & mstr��ˮ�� & "'," & IIf(str��Ŀ���� = "����", 2, 1) & ",'" & _
            str��ϸ���� & "','" & str��Ŀ���� & "','" & str��� & "'," & Format(rs������ϸ!ʵ�ս�� / rs������ϸ!����, "#.###") & "," & _
            rs������ϸ!���� & "," & str�Ը����� & ",'" & UserInfo.���� & "')"
                    
        WriteInfo "������ϸ(д������ϸ):" & strSql
        gcn��Ҧ.Execute strSql
        iLoop = iLoop + 1
        
        gstrSQL = "ZL_����֧����Ŀ_Modify(" & rs������ϸ!�շ�ϸĿID & "," & gintInsure & ",NULL,'" & str��ϸ���� & "','" & _
            str��Ŀ���� & "','" & strҽ������ & "',1)"
        WriteInfo "�޸ı���֧����Ŀ:" & gstrSQL
        Call ExecuteProcedure(gstrSysName)
        
        rs������ϸ.MoveNext
    Wend
    WriteInfo " "
'
'    gstrInput��Ҧ = "$$" & mcur��ҽ�� & "~1~" & mstr��ˮ�� & "~" & gstrIC���� & "~0000$$"
'    gstrOutput��Ҧ = Space(4000)
'    WriteInfo "Ԥ�������:f_Apply(29, " & CDbl(mstr��ˮ��) & ", """ & Replace(gstrInput��Ҧ, String(1053, "0"), "") & """, "" "")"
'    glngReturn = f_Apply(29, CDbl(mstr��ˮ��), gstrInput��Ҧ, gstrOutput��Ҧ)
'    WriteInfo "Ԥ���㷵��:" & gstrOutput��Ҧ
'    �����������_��Ҧ = CheckReturn_��Ҧ()
'    WriteInfo "���Ԥ����"
    �����������_��Ҧ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    �����������_��Ҧ = False
End Function

Public Function �������_��Ҧ(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ�
'        ���������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ����
'        ����һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim lng����ID  As Long, rsTemp As New ADODB.Recordset, datCurr As Date, cur���� As Currency
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, _
        cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency, strTemp As String
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select * From ���˷��ü�¼ Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    lng����ID = rsTemp!����ID
    While Not rsTemp.EOF
        cur���� = cur���� + rsTemp!ʵ�ս��
        rsTemp.MoveNext
    Wend
    
'    gstrOutput��Ҧ = Space(4000)
'    gstrInput��Ҧ = "$$1~" & mcur��ҽ�� & "~1~" & mstr��ˮ�� & "~" & gstrIC���� & "$$"
'    WriteInfo vbCrLf & "�������:f_Apply(30, " & CDbl(mstr��ˮ��) & ", """ & Replace(gstrInput��Ҧ, String(1053, "0"), "") & """, "" "")"
'    glngReturn = f_Apply(30, CDbl(mstr��ˮ��), gstrInput��Ҧ, gstrOutput��Ҧ)
'    WriteInfo "���㷵��:" & gstrOutput��Ҧ
'    �������_��Ҧ = CheckReturn_��Ҧ()
'    If �������_��Ҧ = False Then
'        Exit Function
'    End If
'    strTemp = Split(gstrOutput��Ҧ, "$$")(2)
'    cur���� = CCur(Split(strTemp, "~")(0))
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & _
            lng����ID & "," & Year(datCurr) & ",0,0,0,0," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & _
            cur���� & "," & curȫ�Ը� & ",0,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,'" & mstr��ˮ�� & "')"
    Call ExecuteProcedure(gstrSysName)
    '---------------------------------------------------------------------------------------------
    �������_��Ҧ = True
    WriteInfo "�������"
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ����������_��Ҧ(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, str������ As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, lng����ID As Long, strTemp As String
    Dim datCurr As Date, strSql As String


    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
'    gstrIC���� = makeICInfo(lng����id)
'    If gstrIC���� = "" Then Exit Function
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B" & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    lng����ID = rsTemp("����ID")
    
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
    strTemp = rsTemp!�������ý��
    str������ = rsTemp!��ע
'    strSql = "Insert Into hi_ClinicRx (BillID,DateDiagnose,ChargeType,HospitalID,PIN,ClinicSerial,Department,DepartmentID," & _
'        "Doctor,Disease,DiseaseID,Description,DateOccur,Operator) values ('" & mstr��ˮ�� & "','" & Format(datCurr, "yyyy-mm-dd HH:MM:SS") & _
'        "',1," & Trim(gstrҽԺ����) & ",'" & strҽ���� & "','" & lng����id & "',Null,Null,'" & rs������ϸ!������ & _
'        "','" & str���� & "','" & str����ID & "',Null,'" & Format(datCurr, "yyyy-mm-dd HH:MM:SS") & _
'        "','" & UserInfo.���� & "')"
'    strSql = "Insert Into hi_ClinicPrescription (SerialNum,HospitalID,BillID,DateDiagnose,RecipeSerial,Class,ItemID,ItemName," & _
'        "Specification,Price,Dosage,ScaleSelf,Operator) Values (" & iLoop + lng��ˮ & "," & Trim(gstrҽԺ����) & ",'" & mstrBillNo & _
'        "','" & Format(datCurr, "yyyy-mm-dd HH:MM:SS") & "','" & mstr��ˮ�� & "'," & IIf(str��Ŀ���� = "����", 2, 1) & ",'" & _
'        str��ϸ���� & "','" & str��Ŀ���� & "','" & str��� & "'," & Format(rs������ϸ!ʵ�ս�� / rs������ϸ!����, "#.###") & "," & _
'        rs������ϸ!���� & "," & str�Ը����� & ",'" & UserInfo.���� & "')"
    strSql = "Select * From hi_ClinicRx Where BillID='" & str������ & "'"
    Set rsTemp = gcn��Ҧ.Execute(strSql)
    If rsTemp.EOF Then
        MsgBox "ǰ�û���δ�ҵ��õ������ݣ����ϴ������ݲ����˷�", vbInformation, gstrSysName
        ����������_��Ҧ = False
        Exit Function
    End If
    
    strSql = "Select * From hi_ClinicPrescription Where RecipeSerial='" & str������ & "'"
    Set rsTemp = gcn��Ҧ.Execute(strSql)
    If rsTemp.EOF Then
        MsgBox "ǰ�û���δ�ҵ��õ������ݣ����ϴ������ݲ����˷�", vbInformation, gstrSysName
        ����������_��Ҧ = False
        Exit Function
    End If
    gcn��Ҧ.Execute "Delete hi_ClinicRx Where BillID='" & str������ & "'"
    gcn��Ҧ.Execute "Delete hi_ClinicPrescription Where RecipeSerial='" & str������ & "'"
    
'    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(31)
'
'    '���ýӿ�������
'    gstrInput��Ҧ = "$$" & str������ & "~" & gstrIC���� & "$$"
'    gstrOutput��Ҧ = Space(4000)
'    glngReturn = f_Apply(31, CDbl(mstr��ˮ��), gstrInput��Ҧ, gstrOutput��Ҧ)
'    ����������_��Ҧ = CheckReturn_��Ҧ()
'    If ����������_��Ҧ = False Then
'        Exit Function
'    End If
'    strTemp = Split(gstrOutput��Ҧ, "$$")(2)
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)

    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
        Year(datCurr) & ",0,0,0,0," & intסԺ�����ۼ� & ",0,0,0,-" & strTemp & ",0,0,0," & _
        "0,0,0,0,NULL,NULL,NULL,'" & mstr��ˮ�� & "')"
    Call ExecuteProcedure(gstrSysName)
    ����������_��Ҧ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_��Ҧ(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim strSql As String, strInNote As String, rsTemp As New ADODB.Recordset, str���� As String, str���ֱ��� As String
    Dim rsTmp As New ADODB.Recordset, str������ As String, datCurr As Date, strTemp As String
    Dim lng����ID As Long
    
    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID)   '��Ժ���
'    If rsTmp.BOF Then ��Ժ�Ǽ�_��Ҧ = False: Exit Function
    'ǿ��ѡ����
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & gintInsure
    
    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ������")
    If rsTemp.State = 1 Then
        lng����ID = rsTemp("ID")
        str���� = rsTemp!����
        str���ֱ��� = rsTemp!ID
    Else
        ��Ժ�Ǽ�_��Ҧ = False
        Exit Function
    End If
    
    gstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,A.��Ժ����,A.סԺҽʦ,C.����," & _
            "C.����,D.���� As ���ұ��� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = " & lng��ҳID & " And A.����ID = " & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(32)
    gstrIC���� = makeICInfo(lng����ID)
    
    gstrInput��Ҧ = "$$" & gstrIC���� & "~" & mstr��ˮ�� & "~" & _
        Format(Nvl(rsTemp(0), datCurr), "yyyy-mm-dd") & "~" & Nvl(rsTemp(4), " ") & "~" & strInNote & "~" & _
        str���ֱ��� & "~" & Nvl(rsTemp!סԺ����, " ") & "~" & Nvl(rsTemp!���ұ���, "0") & "~" & Nvl(rsTemp!��Ժ����, "0") & "$$"
    gstrOutput��Ҧ = Space(4000)
    glngReturn = f_Apply(32, CDbl(mstr��ˮ��), gstrInput��Ҧ, gstrOutput��Ҧ)
    ��Ժ�Ǽ�_��Ҧ = CheckReturn_��Ҧ()
    If ��Ժ�Ǽ�_��Ҧ = False Then
        Exit Function
    End If
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ҧ & ",'����ID'," & lng����ID & ")"
    Call ExecuteProcedure(gstrSysName)
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ҧ & ",'˳���'," & mstr��ˮ�� & ")"
    Call ExecuteProcedure(gstrSysName)
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    ��Ժ�Ǽ�_��Ҧ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_��Ҧ = False
End Function

Public Function ���ʴ���_��Ҧ(ByVal str���ݺ� As String, ByVal int���� As Integer, str��Ϣ As String, Optional ByVal lng����ID As Long = 0) As Boolean
    Dim rsTemp As New ADODB.Recordset, lng��ҳID As Long, iLoop As Long, strSql As String, lng��ˮ As Long, _
        rs��ϸ As New ADODB.Recordset, strTemp As String, strסԺ�� As String, str��ϸ���� As String, str��Ŀ���� As String, _
        str��� As String, str��Ŀ���� As String, strҽ������ As String, str�Ը����� As String
    On Error GoTo errHandle
    'ȡ���������ҳID
    gstrSQL = "Select Max(��ҳID) From ������ҳ Where ����id=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    lng��ҳID = rsTemp(0)
    gstrSQL = "Select * From �����ʻ� Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    strסԺ�� = Format(Val(rsTemp!˳���), "0" & String(16, "#")) ' Val(rsTemp!˳���)
    
    'ȡ���˷��ü�¼
    If str���ݺ� <> "" Then
        gstrSQL = "Select * From ���˷��ü�¼ Where ʵ�ս��<>0 And ʵ�ս�� Is Not Null And ��¼״̬<>0 And Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 and ��¼����=" & int���� & " and NO='" & str���ݺ� & "' order by ��ҳID,���"
    Else
        gstrSQL = "Select * From ���˷��ü�¼ Where ʵ�ս��<>0 And ʵ�ս�� Is Not Null And ��¼״̬<>0 And Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 and ����id=" & lng����ID & " And ��ҳid=" & lng��ҳID & " order by ��ҳID,���"
    End If
    Call OpenRecordset(rs��ϸ, gstrSQL)
    
    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(33)
    iLoop = 1
    strSql = "Select Max(SerialNum) From hi_InpatientPrescription"
    Set rsTemp = gcn��Ҧ.Execute(strSql)
    If rsTemp.EOF Then
        lng��ˮ = 0
    Else
        lng��ˮ = Nvl(rsTemp(0), 0)
    End If
    While Not rs��ϸ.EOF
        gstrSQL = "Select ����,����,���,nvl(���,'') as ��� From �շ�ϸĿ Where ID=" & rs��ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, gstrSysName)
        str��ϸ���� = rsTemp!����: str��Ŀ���� = rsTemp!����
        str��� = Left(Left(rsTemp!��� & " |", InStr(rsTemp!��� & " |", "|") - 1), InStr(rsTemp!��� & " |", " ") - 1)
        '�ж���Ŀ����
        str��Ŀ���� = IIf(rsTemp!��� = "5" Or rsTemp!��� = "6", "ҩƷ", IIf(rsTemp!��� = "7", "��ҩ", "����"))
        
        '�ӱ���֧����Ŀ�в����Ƿ��и�ҽ����Ŀ
        gstrSQL = "Select ��Ŀ����,��Ŀ���� From ����֧����Ŀ Where �Ƿ�ҽ��=1 And ����=" & gintInsure & " And �շ�ϸĿID=" & rs��ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, gstrSysName)
        If rsTemp.EOF Then      'û����Ŀ����
            mcur��ҽ�� = mcur��ҽ�� + rs��ϸ!ʵ�ս��
            If str��Ŀ���� = "ҩƷ" Then    '����ΪҩƷʱ��ҽ������Ϊ�����ࡱ
                strҽ������ = "����": str�Ը����� = "1"
            Else        '��Ŀ����Ϊ���ƻ���ҩʱ��ҽ������Ϊ�����ࡱ
                strҽ������ = "����": str�Ը����� = "0"
            End If
        Else            '�и���Ŀʱ����
            str��ϸ���� = rsTemp!��Ŀ����
            If str��Ŀ���� = "����" Then
                gstrSQL = "Select * From hi_Diagnose Where DiagnoseID='" & str��ϸ���� & "'"
            Else
                gstrSQL = "Select * From hi_Medicine Where MedicineID='" & str��ϸ���� & "'"
            End If
            Set rsTemp = gcn��Ҧ.Execute(gstrSQL)
            If rsTemp.EOF Then      '���ҽ������Ŀ¼��δ�ҵ�����Ŀ
                If str��Ŀ���� = "ҩƷ" Then    '����ΪҩƷʱ��ҽ������Ϊ�����ࡱ
                    strҽ������ = "����": str�Ը����� = "1"
                Else        '��Ŀ����Ϊ���ƻ���ҩʱ��ҽ������Ϊ�����ࡱ
                    strҽ������ = "����": str�Ը����� = "0"
                End If
            Else        '���ҽ������Ŀ¼���и�ҩƷ
                strҽ������ = IIf(rsTemp!zfbl = 0, "����", IIf(rsTemp!zfbl = 1, "����", "����"))
                str�Ը����� = rsTemp!zfbl
            End If
        End If
        strSql = "Insert Into hi_InpatientPrescription (SerialNum,InpatientID,HospitalID,FeeType,RecipeSerial,DateDiagnose," & _
            "Class,ItemID,ItemName,Specification,Price,Dosage,ScaleSelf,Operator) Values (" & iLoop + lng��ˮ & ",'" & _
            strסԺ�� & "'," & Trim(gstrҽԺ����) & ",1,Null,'" & Format(rs��ϸ!����ʱ��, "yyyy-mm-dd HH:MM:SS") & _
            "'," & IIf(str��Ŀ���� = "����", 1, 2) & ",'" & str��ϸ���� & _
            "','" & str��Ŀ���� & "','" & str��� & "'," & Format(Nvl(rs��ϸ!ʵ�ս��, 0) / (rs��ϸ!���� * rs��ϸ!����), _
            "0.000") & "," & rs��ϸ!���� * rs��ϸ!���� & "," & str�Ը����� & ",'" & UserInfo.���� & "')"
                    
        WriteInfo "������ϸ(д������ϸ):" & strSql
        gcn��Ҧ.Execute strSql
        iLoop = iLoop + 1
        
        gstrSQL = "ZL_����֧����Ŀ_Modify(" & rs��ϸ!�շ�ϸĿID & "," & gintInsure & ",NULL,'" & str��ϸ���� & "','" & _
            str��Ŀ���� & "','" & strҽ������ & "',1)"
        WriteInfo "�޸ı���֧����Ŀ:" & gstrSQL
        Call ExecuteProcedure(gstrSysName)
        
        gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rs��ϸ!ID & "')"
        Call ExecuteProcedure(gstrSysName)
        rs��ϸ.MoveNext
    Wend
    '���ýӿ�
'    gstrIC���� = makeICInfo(lng����id)
'    gstrInput��Ҧ = "$$" & strסԺ�� & "~" & gstrIC���� & "~0000$$"
'    gstrOutput��Ҧ = Space(4000)
'    glngReturn = f_Apply(33, CDbl(mstr��ˮ��), gstrInput��Ҧ, gstrOutput��Ҧ)
'    ���ʴ���_��Ҧ = CheckReturn_��Ҧ()
'    If ���ʴ���_��Ҧ = False Then Exit Function
    ���ʴ���_��Ҧ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ���ʴ���_��Ҧ = False
End Function

Public Function סԺ�������_��Ҧ(rs������ϸ As Recordset, lng����ID As Long, strҽ���� As String) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim datCurr As Date, strסԺ�� As String
    
    On Error GoTo errHandle
    Set rs��ϸ = rs������ϸ.Clone
    If rs��ϸ.EOF = True Then
        MsgBox "û����д�շѼ�¼", vbExclamation, gstrSysName
        Exit Function
    End If
    '��Ҫ���ϴ�������ϸ
    If ���ʴ���_��Ҧ("", 0, "", lng����ID) = False Then Exit Function
    
    gstrSQL = "Select * From �����ʻ� Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    strסԺ�� = Format(Val(rsTemp!˳���), "0" & String(16, "#")) ' Val(rsTemp!˳���)
    
    '�����ҽ�����
'    mcur��ҽ�� = 0
'    While Not rs��ϸ.EOF
'        gstrSQL = "Select A.���,B.��Ŀ����,B.��Ŀ����,Nvl(A.���,"") As ��� From ����֧����Ŀ A,�շ�ϸĿ B " & _
'            "Where A.ID=B.�շ�ϸĿID And B.�Ƿ�ҽ��=1 And B.����=" & gintInsure & " And A.ID=" & rs��ϸ!�շ�ϸĿID
'        Call OpenRecordset(rsTemp, gstrSysName)
'        If Not rsTemp.EOF Then
'            '�ж�ҽ��ǰ�û����Ƿ��и���Ŀ
'            If rsTemp(0) = "6" Or rsTemp(0) = "7" Or rsTemp(0) = "5" Then
'                gstrSQL = "Select * From hi_Medicine Where MedicineID='" & rsTemp(1) & "'"
'            Else
'                gstrSQL = "Select * From hi_Diagnose Where DiagnoseID='" & rsTemp(1) & "'"
'            End If
'            Set rsTemp = gcn��Ҧ.Execute(gstrSQL)
'            If rsTemp.EOF Then mcur��ҽ�� = mcur��ҽ�� + rs��ϸ!ʵ�ս��
'        Else
'            mcur��ҽ�� = mcur��ҽ�� + rs��ϸ!ʵ�ս��
'        End If
'        rs��ϸ.MoveNext
'    Wend
'    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(34)
'    '���ýӿ�
'    gstrIC���� = makeICInfo(lng����id)
'    gstrInput��Ҧ = "$$" & strסԺ�� & "~" & mcur��ҽ�� & "~" & gstrIC���� & "~0000$$"
'    gstrOutput��Ҧ = Space(4000)
'    glngReturn = f_Apply(34, CDbl(mstr��ˮ��), gstrInput��Ҧ, gstrOutput��Ҧ)
'    Call CheckReturn_��Ҧ
    סԺ�������_��Ҧ = "ҽ������;0;0"
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    סԺ�������_��Ҧ = ""
End Function

Public Function ��Ժ�Ǽ�_��Ҧ(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    '����״̬���޸�
    Dim rsTemp As New ADODB.Recordset, datCurr As Date, bln����ó�Ժ As Boolean, strסԺ�� As String, _
        strInNote As String, str���ֱ��� As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    '���ô�סԺ�Ƿ�û�з��÷���
    gstrSQL = "Select nvl(sum(ʵ�ս��),0) as ��� from ���˷��ü�¼ where nvl(���ӱ�־,0)<>9 and ����ID=" & lng����ID & " and ��ҳID=" & lng��ҳID
    Call OpenRecordset(rsTemp, "���˳�Ժ")
    If rsTemp.EOF = True Then
        bln����ó�Ժ = True
    Else
        bln����ó�Ժ = (rsTemp("���") = 0)
    End If
    
    gstrSQL = "Select * From �����ʻ� Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    strסԺ�� = Format(Val(rsTemp!˳���), "0" & String(16, "#")) ' Val(rsTemp!˳���)
    
    If bln����ó�Ժ = True Then
        '������Ժ�Ǽǳ���
        mstr��ˮ�� = ���뽻����ˮ_��Ҧ(40)
        gstrIC���� = makeICInfo(lng����ID)
        
        '���ýӿ�
        gstrInput��Ҧ = "$$" & strסԺ�� & "~" & strסԺ�� & "~" & gstrIC���� & "$$"
        gstrOutput��Ҧ = Space(4000)
        glngReturn = f_Apply(40, CDbl(mstr��ˮ��), gstrInput��Ҧ, gstrOutput��Ҧ)
        ��Ժ�Ǽ�_��Ҧ = CheckReturn_��Ҧ()
        Exit Function
    End If
    
    '��ȡ��Ժ���
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, True, True)
    
    'ǿ��ѡ����
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & gintInsure
    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ȷ�Ｒ��")
    If rsTemp.State = 1 Then
        str���ֱ��� = rsTemp!ID
    Else
        ��Ժ�Ǽ�_��Ҧ = False
        Exit Function
    End If
    '��ȡסԺҽʦ
    gstrSQL = "select סԺҽʦ from ������ҳ Where ��ҳID = " & lng��ҳID & " And ����ID = " & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "����ȡ�ò��˵���Ժ�Ǽ���Ϣ", vbInformation, gstrSysName
        Exit Function
    End If
    
    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(35)
    gstrIC���� = makeICInfo(lng����ID)
    
    '���ýӿ�
    gstrInput��Ҧ = "$$" & strסԺ�� & "~" & Nvl(rsTemp(0), " ") & "~" & strInNote & "~" & _
        str���ֱ��� & "~" & Format(datCurr, "yyyy-mm-dd") & "$$"
    gstrOutput��Ҧ = Space(4000)
    glngReturn = f_Apply(35, CDbl(mstr��ˮ��), gstrInput��Ҧ, gstrOutput��Ҧ)
    ��Ժ�Ǽ�_��Ҧ = CheckReturn_��Ҧ()
    If ��Ժ�Ǽ�_��Ҧ = False Then Exit Function
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    ��Ժ�Ǽ�_��Ҧ = False
End Function

Public Function סԺ����_��Ҧ(lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ�
'        ���������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ����
'        ����һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim lng����ID  As Long, rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str����Ա As String, datCurr As Date, str������ As String, strTemp As String
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency
    Dim cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    
    On Error GoTo errHandle
    
    gstrSQL = "Select * From ���˷��ü�¼ Where ��¼״̬<>0 And nvl(���ӱ�־,0)<>9 and ����ID=" & lng����ID
    Call OpenRecordset(rs��ϸ, gstrSysName)
    If rs��ϸ.EOF Then
        MsgBox "û�з�����ϸ�����ܽ��г�Ժ����", vbInformation, gstrSysName
        Exit Function
    End If
    lng����ID = rs��ϸ!����ID
    
    gstrSQL = "Select nvl(˳���,0) as ˳��� From �����ʻ� Where ����ID=" & lng����ID & " And ����=" & TYPE_��Ҧ
    Call OpenRecordset(rsTemp, gstrSysName)
    str������ = Format(Val(rsTemp!˳���), "0" & String(16, "#")) ' Nvl(rsTemp!˳���)
    datCurr = zlDatabase.Currentdate
    
'    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(36)
'    gstrIC���� = makeICInfo(lng����id)
    
    '�����ҽ����Ŀ���
    mcur��ҽ�� = 0
    While Not rs��ϸ.EOF
'        gstrSQL = "Select A.���,B.��Ŀ����,B.��Ŀ����,Nvl(A.���,"") As ��� From ����֧����Ŀ A,�շ�ϸĿ B " & _
'            "Where A.ID=B.�շ�ϸĿID And B.�Ƿ�ҽ��=1 And B.����=" & gintInsure & " And A.ID=" & rs��ϸ!�շ�ϸĿID
'        Call OpenRecordset(rsTemp, gstrSysName)
'        If Not rsTemp.EOF Then
'            '�ж�ҽ��ǰ�û����Ƿ��и���Ŀ
'            If rsTemp(0) = "6" Or rsTemp(0) = "7" Or rsTemp(0) = "5" Then
'                gstrSQL = "Select * From hi_Medicine Where MedicineID='" & rsTemp(1) & "'"
'            Else
'                gstrSQL = "Select * From hi_Diagnose Where DiagnoseID='" & rsTemp(1) & "'"
'            End If
'            Set rsTemp = gcn��Ҧ.Execute(gstrSQL)
'            If rsTemp.EOF Then mcur��ҽ�� = mcur��ҽ�� + rs��ϸ!ʵ�ս��
'        Else
            mcur��ҽ�� = mcur��ҽ�� + Nvl(rs��ϸ!ʵ�ս��, 0)
'        End If
        rs��ϸ.MoveNext
    Wend
'    gstrInput��Ҧ = "$$1~" & str������ & "~" & mcur��ҽ�� & "~" & gstrIC���� & "$$"
'    gstrOutput��Ҧ = Space(4000)
'    glngReturn = f_Apply(36, CDbl(mstr��ˮ��), gstrInput��Ҧ, gstrOutput��Ҧ)
'    סԺ����_��Ҧ = CheckReturn_��Ҧ()
'    If סԺ����_��Ҧ = False Then Exit Function
'    strTemp = Split(gstrOutput��Ҧ, "$$")(2)
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)

    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & _
            lng����ID & "," & Year(datCurr) & ",0," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,NULL,NULL," & mcur��ҽ�� & _
            ",0,0,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,'" & _
            str������ & "~" & mstr��ˮ�� & "')"
    Call ExecuteProcedure(gstrSysName)
    סԺ����_��Ҧ = True
    '---------------------------------------------------------------------------------------------

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_��Ҧ(lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, lng����ID As Long, str��ˮ�� As String, str������ As String, _
        lng����ID As Long
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, _
        curͳ�ﱨ���ۼ� As Currency, intסԺ�����ۼ� As Integer, curƱ���ܽ�� As Currency
    Dim datCurr As Date
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ�� From ���˷��ü�¼ Where ��¼״̬<>0 And nvl(���ӱ�־,0)<>9 and ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "û�ҵ����˵ķ�����ϸ��¼�������˷�", vbInformation, gstrSysName
        Exit Function
    End If
    lng����ID = rsTemp("����ID")
    Do Until rsTemp.EOF
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B" & _
              " where b.nvl(���ӱ�־,0)<>9 and a.nvl(���ӱ�־,0)<>9 and A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    lng����ID = rsTemp("����ID")
    
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=" & gintInsure & " and ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    If rsTemp.EOF = True Then
        MsgBox "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If IsNull(rsTemp!��ע) Then
        MsgBox "�õ��ݵľ����Ŷ�ʧ���������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    str������ = Split(rsTemp!��ע, "~")(0)
    str��ˮ�� = Split(rsTemp!��ע, "~")(1)
    
    '���ýӿ�������
'    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(37)
'    gstrIC���� = makeICInfo(lng����id)
'
'    '���ýӿ�
'    gstrInput��Ҧ = "$$" & str��ˮ�� & "~" & gstrIC���� & "$$"
'    gstrOutput��Ҧ = Space(4000)
'    glngReturn = f_Apply(37, CDbl(mstr��ˮ��), gstrInput��Ҧ, gstrOutput��Ҧ)
'    סԺ�������_��Ҧ = CheckReturn_��Ҧ()
'    If סԺ�������_��Ҧ = False Then Exit Function
'    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����id & "," & TYPE_��Ҧ & ",'˳���'," & str��ˮ�� & ")"
'    Call ExecuteProcedure(gstrSysName)
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0,0,0,0,0,0," & _
        "NULL,NULL,NULL,'" & str������ & "~" & str��ˮ�� & "')"
    Call ExecuteProcedure(gstrSysName)

    סԺ�������_��Ҧ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��ݱ�ʶ_��Ҧ(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim frmIDentified As New frmIdentify��Ҧ
    Dim strPatiInfo As String, cur��� As Currency, str������ As String
    Dim arr, datCurr As Date, str����� As String
    Dim strSql As String, rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    'MODIFIED BY ZYB ����ҽ���ӿڿ���
    strPatiInfo = frmIDentified.GetPatient(bytType, lng����ID)
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
        ��ݱ�ʶ_��Ҧ = ""
        MsgBox "ҽ��������Ϣ��ȡʧ��", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    ��ݱ�ʶ_��Ҧ = strPatiInfo
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_��Ҧ = ""
End Function

Public Function �������_��Ҧ(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    '��Ҧ������ȡ�����ʻ����
    �������_��Ҧ = 0
End Function

Public Function ת��ת��_��Ҧ(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim strSql As String, strInNote As String, rsTemp As New ADODB.Recordset, str���� As String, str���ֱ��� As String
    Dim rsTmp As New ADODB.Recordset, str������ As String, datCurr As Date, strTemp As String
    Dim lng����ID As Long
    
    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID)   '��Ժ���
    If rsTmp.BOF Then ת��ת��_��Ҧ = False: Exit Function
    'ǿ��ѡ����
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & gintInsure
    
    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ������")
    If rsTemp.State = 1 Then
        lng����ID = rsTemp("ID")
        str���� = rsTemp!����
        str���ֱ��� = rsTemp!ID
    Else
        ת��ת��_��Ҧ = False
        Exit Function
    End If
    
    gstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,A.��Ժ����,A.סԺҽʦ,C.����," & _
            "C.����,D.���� As ���ұ���,C.˳��� As סԺ��ˮ from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = " & lng��ҳID & " And A.����ID = " & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(38)
    gstrIC���� = makeICInfo(lng����ID)
    
    gstrInput��Ҧ = "$$" & rsTemp!סԺ��ˮ & "~" & Format(datCurr, "yyyy-mm-dd") & "~" & _
        rsTemp(3) & "~" & Nvl(rsTemp(4), " ") & "~" & strInNote & "~" & _
        str���ֱ��� & "~" & Nvl(rsTemp!סԺ����, " ") & "~" & Nvl(rsTemp!���ұ���, "0") & "$$"
    gstrOutput��Ҧ = Space(4000)
    glngReturn = f_Apply(38, CDbl(mstr��ˮ��), gstrInput��Ҧ, gstrOutput��Ҧ)
    ת��ת��_��Ҧ = CheckReturn_��Ҧ()
    If ת��ת��_��Ҧ = False Then
        Exit Function
    End If
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ҧ & ",'����ID'," & lng����ID & ")"
    Call ExecuteProcedure(gstrSysName)
    ת��ת��_��Ҧ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ת��ת��_��Ҧ = False
End Function

