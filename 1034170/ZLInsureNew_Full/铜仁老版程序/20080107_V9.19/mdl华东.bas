Attribute VB_Name = "mdl����"
Option Explicit
Private mcurͳ���� As Currency, mcur����֧�� As Currency
Public gcn���� As New ADODB.Connection, mstrSavePath As String

Public Const MAX_PATH = 260

Public Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Public Function BrowPath(lWindowHwnd As Long, Optional ByVal sTitle As String = "") As String
    Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    With udtBI
        '�����������
        .hwndOwner = lWindowHwnd
        '����ѡ�е�Ŀ¼
        .ulFlags = BIF_RETURNONLYFSDIRS
        If sTitle = "" Then
            .lpszTitle = "��ѡ����ʼ�������ļ��У�"
        Else
            .lpszTitle = sTitle
        End If
    End With
    
    '�����������
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        '��ȡ·��
        SHGetPathFromIDList lpIDList, sPath
        '�ͷ��ڴ�
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    BrowPath = sPath
End Function


Public Function ҽ����ʼ��_����() As Boolean
'���ܣ������Ƿ�������ӵ�ǰ�÷�������
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSql As String, rs���� As New ADODB.Recordset
    '��������Ѿ��򿪣��ǾͲ����ٲ���
    If gcn����.State = adStateOpen Then
        ҽ����ʼ��_���� = True
        Exit Function
    End If
     
    On Error GoTo errH
    
    '���ȶ���������������
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        If rsTemp!������ = "�ļ����λ��" Then mstrSavePath = rsTemp!����ֵ
        rsTemp.MoveNext
    Loop
    If Trim(mstrSavePath) = "" Then
        MsgBox "�뵽ҽ�����������������ļ����λ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error Resume Next
    gcn����.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=Visual FoxPro Tables;UID=;SourceDB=" & mstrSavePath & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"""
    gcn����.CursorLocation = adUseClient
    gcn����.Open
    
    If Err <> 0 Then
        MsgBox "�ļ����λ��ָ������", vbInformation, gstrSysName
        ҽ����ʼ��_���� = False
        Exit Function
    End If
    ҽ����ʼ��_���� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    ҽ����ʼ��_���� = False
End Function

Public Function ҽ������_����() As Boolean
    ҽ������_���� = frmSet����.ShowMe(gintInsure)
End Function

Public Function �������_����(lng����id As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select * From �����ʻ� Where ����id=" & lng����id & " And ����=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    �������_���� = Nvl(rsTemp!�ʻ����, 0)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte = 0, Optional lng����id As Long = 0) As String
    '����ҽ��û�ṩר�ŵ������֤�ӿڣ�ͨ����ȡ�Һŵ�����ʵ����֤
    Dim strTemp As String
    strTemp = frmIdentify����.Identify(bytType, lng����id)
    Unload frmIdentify����
    If strTemp = "" Then
        MsgBox "δ��ȡ������Ϣ", vbInformation, gstrSysName
    Else
        ��ݱ�ʶ_���� = strTemp
    End If
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
'��Ϊ����δ�ṩԤ����ӿڣ�������õ��Ľ�������Ϊҽ���������ʽ���ݣ����õ�����ʱҽ������ʽ����
    Dim str��ˮ�� As String, lng����id As Long, datCurr As Date, strSql As String
    Dim rsTemp As New ADODB.Recordset, rsDBF As New ADODB.Recordset, lng��� As Long
    Dim strCardNO As String
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
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�в��˷��ã����ܽ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    datCurr = zlDatabase.Currentdate
    lng����id = rs��ϸ(0)
    gstrSQL = "Select ���� From �����ʻ� Where ����id=" & lng����id & " And ����=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "û���ҵ�������Ϣ��ҽ��ѡ�����", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!����
    '������ˮ��
    str��ˮ�� = Mid(Format(datCurr, "YYMMDDHHMMSS"), 2, 10) & Format(lng����id, "0####")
    
    '�ж��Ƿ���ҽ������δ��Ӧ
    Do Until rs��ϸ.EOF
        gstrSQL = "select A.��Ŀ����,B.���� from (select * from ����֧����Ŀ where ����=" & gintInsure & ") A, �շ�ϸĿ B where A.�շ�ϸĿid(+)=B.id and B.id = " & rs��ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, gstrSysName)
        If IsNull((rsTemp!��Ŀ����)) Then
            MsgBox "<" & rsTemp!���� & ">δ��Ӧҽ������,���Ƚ��ж���", vbInformation, gstrSysName
            Exit Function
        End If
        rs��ϸ.MoveNext
    Loop
    
    '����DBF�ļ�
    On Error Resume Next
    gcn����.Execute "Drop Table " & mstrSavePath & "\YM" & str��ˮ��
    
    On Error GoTo errHandle
    gcn����.Execute "Create Table " & mstrSavePath & "\YM" & str��ˮ�� & " (IDNo C(18),CaseNo C(15),OrderNo N(18,4)," & _
        "IntelCode C(14),CName C(70),SubCode C(8),Standard C(20),CUnit C(4),Num N(18,4),Price N(18,4),SumJe N(18,4)," & _
        "SelfJe N(18,4))"
    lng��� = 1
    rs��ϸ.MoveFirst
    While Not rs��ϸ.EOF
        gstrSQL = "Select A.��Ŀ����,B.ID,B.����,B.���,B.���㵥λ From ����֧����Ŀ A,�շ�ϸĿ B Where B.ID=A.�շ�ϸĿID And A.�շ�ϸĿid=" & rs��ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, gstrSysName)             '��Ϊ֮ǰ������Ƿ���ж��룬���Զ����ļ�¼һ�������
        
        '���š���ˮ�š���š����롢���ơ���Ŀ���롢��񡢼�����λ�����������ۡ����Էѽ��
        gcn����.Execute "Insert Into " & mstrSavePath & "\YM" & str��ˮ�� & " values ('','" & str��ˮ�� & "'," & _
            lng��� & ",'" & Trim(rsTemp!��Ŀ����) & "','" & Trim(rsTemp!����) & "','" & Trim(rsTemp!���) & "','" & Trim(rsTemp!���㵥λ) & "'," & _
            "''," & rs��ϸ!���� & "," & rs��ϸ!���� & "," & rs��ϸ!ʵ�ս�� & "," & _
            "0)"
        lng��� = lng��� + 1
        rs��ϸ.MoveNext
    Wend
    On Error GoTo errHandle
    '�ȴ����ؽ�������
    If frm�ȴ����ػ���.waitReturn(mstrSavePath & "\SM" & str��ˮ��) = False Then
        MsgBox "Ԥ���㱻��ֹ", vbInformation, gstrSysName
        On Error Resume Next
        gcn����.Execute "Drop Table " & mstrSavePath & "\YM" & str��ˮ��
        Unload frm�ȴ����ػ���
        Exit Function
    End If
    Unload frm�ȴ����ػ���
    
    '���ؽ�����
    strSql = "Select * From " & mstrSavePath & "\SM" & str��ˮ��
    Set rsTemp = gcn����.Execute(strSql)
    mcur����֧�� = Val(rsTemp!JkAccR)
    mcurͳ���� = Val(rsTemp!JkSocialR)
    str���㷽ʽ = "�����ʻ�;" & Val(rsTemp!JkAccR) & ";0"
    str���㷽ʽ = str���㷽ʽ & "|ͳ�����;" & Val(rsTemp!JkSocialR) & ";0"
    On Error Resume Next
    gcn����.Execute "Drop Table " & mstrSavePath & "\YM" & str��ˮ��
    gcn����.Execute "Drop Table " & mstrSavePath & "\SM" & str��ˮ��
    �����������_���� = True
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curƱ���ܽ�� As Currency
    Dim datCurr As Date, lng����id As Long
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ�� From ���˷��ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    Do Until rsTemp.EOF
        If lng����id = 0 Then lng����id = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����id, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����id & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + mcur����֧�� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� + mcurͳ���� & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure(gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & lng����id & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + mcur����֧�� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� + mcurͳ���� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� & ",0,0," & _
        "0," & mcurͳ���� & ",0,0," & mcur����֧�� & ",Null,Null,Null,Null)"
    Call ExecuteProcedure(gstrSysName)

    �������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����id As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long, str��ˮ�� As String, str������ As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, sngArrInfo(20) As Single
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curƱ���ܽ�� As Currency, lngErr As Long
    Dim datCurr As Date, strRecCode As String, strBillCode As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ�� From ���˷��ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    Do Until rsTemp.EOF
        If lng����id = 0 Then lng����id = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B" & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    lng����ID = rsTemp("����ID")
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=" & gintInsure & " and ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF = True Then
        MsgBox "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        ����������_���� = False
        Exit Function
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����id, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����id & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - Nvl(rsTemp("�����ʻ�֧��"), 0) & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure(gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & lng����id & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - Nvl(rsTemp("�����ʻ�֧��"), 0) & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        Nvl(rsTemp("�����ʻ�֧��"), 0) * -1 & ",Null,Null,Null,Null)"
    Call ExecuteProcedure(gstrSysName)

    ����������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����(rs��ϸ As ADODB.Recordset, lng����id As Long, strҽ���� As String) As String
'��Ϊ����δ�ṩԤ����ӿڣ�������õ��Ľ�������Ϊҽ���������ʽ���ݣ����õ�����ʱҽ������ʽ����
    Dim str��ˮ�� As String, datCurr As Date, strSql As String
    Dim rsTemp As New ADODB.Recordset, rsDBF As New ADODB.Recordset, lng��� As Long
    Dim strCardNO As String
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
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�в��˷��ã����ܽ���", vbInformation, gstrSysName
        Exit Function
    End If

    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select ���� From �����ʻ� Where ����id=" & lng����id & " And ����=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "û���ҵ�������Ϣ��ҽ��ѡ�����", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!����
    '������ˮ��
    gstrSQL = "Select Max(��ҳID) From ������ҳ Where ����id=" & lng����id
    Call OpenRecordset(rsTemp, gstrSysName)
    str��ˮ�� = Format(lng����id, "0######") & "_" & rsTemp(0)

    '�ж��Ƿ���ҽ������δ��Ӧ
    Do Until rs��ϸ.EOF
        gstrSQL = "select A.��Ŀ����,B.���� from (select * from ����֧����Ŀ where ����=" & gintInsure & ") A, �շ�ϸĿ B where A.�շ�ϸĿid(+)=B.id and B.id = " & rs��ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, gstrSysName)
        If IsNull((rsTemp!��Ŀ����)) Then
            MsgBox "<" & rsTemp!���� & ">δ��Ӧҽ������,���Ƚ��ж���", vbInformation, gstrSysName
            Exit Function
        End If
        rs��ϸ.MoveNext
    Loop
    ���ʴ���_���� "", 0, "", lng����id
    '����DBF�ļ�
'    gcn����.Execute "Create Table " & mstrSavePath & "\YZ" & str��ˮ�� & " (IDNo C(18),CaseNo C(15),OrderNo N(18,4)," & _
'        "IntelCode C(14),CName C(70),SubCode C(8),Standard C(20),CUnit C(4),Num N(18,4),Price N(18,4),SumJe N(18,4)," & _
'        "SelfJe N(18,4))"
'    strSql = "Select * From " & mstrSavePath & "\YZ" & str��ˮ��
'    lng��� = 1
'    rs��ϸ.MoveFirst
'    While Not rs��ϸ.EOF
'        gstrSQL = "Select A.��Ŀ����,B.ID,B.����,B.���,B.���㵥λ From ����֧����Ŀ A,�շ�ϸĿ B Where B.ID=A.�շ�ϸĿID And A.�շ�ϸĿid=" & rs��ϸ!�շ�ϸĿID
'        Call OpenRecordset(rsTemp, gstrSysName)             '��Ϊ֮ǰ������Ƿ���ж��룬���Զ����ļ�¼һ�������
'
'        '���š���ˮ�š���š����롢���ơ���Ŀ���롢��񡢼�����λ�����������ۡ����Էѽ��
'        gcn����.Execute "Insert Into " & mstrSavePath & "\YZ" & str��ˮ�� & " values ('" & strCardNO & "','" & str��ˮ�� & "'," & _
'            lng��� & ",'" & Trim(rsTemp!��Ŀ����) & "','" & Trim(rsTemp!����) & "','" & Trim(rsTemp!���) & "','" & Trim(rsTemp!���㵥λ) & "'," & _
'            "''," & rs��ϸ!���� & "," & rs��ϸ!���� & "," & rs��ϸ!ʵ�ս�� & "," & _
'            "0)"
'        lng��� = lng��� + 1
'        rs��ϸ.MoveNext
'    Wend
    On Error GoTo errHandle
    
    '�ȴ����ؽ�������
    If frm�ȴ����ػ���.waitReturn(mstrSavePath & "\SZ" & str��ˮ��) = False Then
        MsgBox "Ԥ���㱻��ֹ", vbInformation, gstrSysName
'        gcn����.Execute "Drop Table " & mstrSavePath & "\YZ" & str��ˮ��
        Unload frm�ȴ����ػ���
        Exit Function
    End If
    Unload frm�ȴ����ػ���
    
    '���ؽ�����
    strSql = "Select Sum(JkaccR) As JkaccR,Sum(JkSocialR) As JkSocialR From " & mstrSavePath & "\SZ" & str��ˮ��
    Set rsTemp = gcn����.Execute(strSql)
    mcur����֧�� = Val(rsTemp!JkAccR)
    mcurͳ���� = Val(rsTemp!JkSocialR)
    סԺ�������_���� = "�����ʻ�;" & Val(rsTemp!JkAccR) & ";0"
    סԺ�������_���� = סԺ�������_���� & "|ͳ�����;" & Val(rsTemp!JkSocialR) & ";0"
    On Error Resume Next
    gcn����.Execute "Drop Table " & mstrSavePath & "\SZ" & str��ˮ��
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long, ByVal lng����id As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curƱ���ܽ�� As Currency
    Dim datCurr As Date
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ�� From ���˷��ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    Do Until rsTemp.EOF
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����id, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����id & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + mcur����֧�� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� + mcurͳ���� & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure(gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & lng����id & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + mcur����֧�� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� + mcurͳ���� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� & ",0,0," & _
        "0," & mcurͳ���� & ",0,0," & mcur����֧�� & ",Null,Null,Null,Null)"
    Call ExecuteProcedure(gstrSysName)

    סԺ����_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����(lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long, str��ˮ�� As String, str������ As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, sngArrInfo(20) As Single
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency, lng����id As Long
    Dim intסԺ�����ۼ� As Integer, curƱ���ܽ�� As Currency, lngErr As Long
    Dim datCurr As Date, strRecCode As String, strBillCode As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ�� From ���˷��ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    Do Until rsTemp.EOF
        If lng����id = 0 Then lng����id = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B" & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    lng����ID = rsTemp("����ID")
    
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=" & gintInsure & " and ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF = True Then
        MsgBox "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        סԺ�������_���� = False
        Exit Function
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����id, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����id & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - Nvl(rsTemp("�����ʻ�֧��"), 0) & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure(gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & lng����id & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - Nvl(rsTemp("�����ʻ�֧��"), 0) & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        Nvl(rsTemp("�����ʻ�֧��"), 0) * -1 & ",Null,Null,Null,Null)"
    Call ExecuteProcedure(gstrSysName)

    סԺ�������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_����(lng����id As Long, lng��ҳID As Long) As Boolean
    On Error GoTo errHandle
    '��HIS֮�еĻ������ݽ����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����id & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_���� = False
End Function

Public Function ��Ժ�Ǽ�_����(lng����id As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    
    On Error GoTo errHandle
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����id & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_���� = False
End Function

Public Function ���ʴ���_����(ByVal str���ݺ� As String, ByVal int���� As Integer, str��Ϣ As String, Optional ByVal lng����id As Long = 0) As Boolean
    Dim rs��ϸ As New ADODB.Recordset, lng��ҳID As Long, rsTemp As New ADODB.Recordset
    Dim str��ˮ�� As String, datCurr As Date, strSql As String, lng��� As Long
    Dim strCardNO As String
    If str���ݺ� <> "" Then
        gstrSQL = "Select ����id From ���˷��ü�¼ Where NO='" & str���ݺ� & "'"
        Call OpenRecordset(rsTemp, gstrSysName)
        lng����id = rsTemp(0)
    End If
    gstrSQL = "Select Max(��ҳID) From ������ҳ Where ����id=" & lng����id
    Call OpenRecordset(rsTemp, gstrSysName)
    lng��ҳID = Nvl(rsTemp(0), 1)
    If str���ݺ� <> "" Then
        gstrSQL = "Select * From ���˷��ü�¼ Where ��¼״̬<>0 And Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 and ��¼����=" & int���� & " and NO='" & str���ݺ� & "' order by ��ҳID,���"
    Else
        gstrSQL = "Select * From ���˷��ü�¼ Where ��¼״̬<>0 And Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 and ����id=" & lng����id & " And ��ҳid=" & lng��ҳID & " And (�Ƿ��ϴ� Is Null Or �Ƿ��ϴ�=0) order by ��ҳID,���"
    End If
    Call OpenRecordset(rs��ϸ, gstrSQL)
    
    On Error GoTo errHandle
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�в��˷��ã����ܽ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "Select * From ���˷��ü�¼ Where nvl(���ӱ�־,0)<>9 and ����id=" & lng����id & " And ��ҳid=" & lng��ҳID & " And �Ƿ��ϴ�=1 order by ��ҳID,���"
    Call OpenRecordset(rsTemp, gstrSysName)
    lng��� = rsTemp.RecordCount + 1
    
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select ���� From �����ʻ� Where ����id=" & lng����id & " And ����=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "û���ҵ�������Ϣ��ҽ��ѡ�����", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!����
    '������ˮ��
    str��ˮ�� = Format(lng����id, "0######") & "_" & lng��ҳID
    
    '�ж��Ƿ���ҽ������δ��Ӧ
    Do Until rs��ϸ.EOF
        gstrSQL = "select A.��Ŀ����,B.���� from (select * from ����֧����Ŀ where ����=" & gintInsure & ") A, �շ�ϸĿ B where A.�շ�ϸĿid(+)=B.id and B.id = " & rs��ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, gstrSysName)
        If IsNull((rsTemp!��Ŀ����)) Then
            MsgBox "<" & rsTemp!���� & ">δ��Ӧҽ������,���Ƚ��ж���", vbInformation, gstrSysName
            Exit Function
        End If
        rs��ϸ.MoveNext
    Loop
    
    '����DBF�ļ�
    On Error Resume Next
    gcn����.Execute "Drop Table " & mstrSavePath & "\YZ" & str��ˮ��
    
    On Error GoTo errHandle
    gcn����.Execute "Create Table " & mstrSavePath & "\YZ" & str��ˮ�� & " (IDNo C(18),CaseNo C(15),OrderNo N(18,4)," & _
        "IntelCode C(14),CName C(70),SubCode C(8),Standard C(20),CUnit C(4),Num N(18,4),Price N(18,4),SumJe N(18,4)," & _
        "SelfJe N(18,4))"
    strSql = "Select * From " & mstrSavePath & "\YZ" & str��ˮ��
    rs��ϸ.MoveFirst
    While Not rs��ϸ.EOF
        gstrSQL = "Select A.��Ŀ����,B.ID,B.����,B.���,B.���㵥λ From ����֧����Ŀ A,�շ�ϸĿ B Where B.ID=A.�շ�ϸĿID And A.�շ�ϸĿid=" & rs��ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, gstrSysName)             '��Ϊ֮ǰ������Ƿ���ж��룬���Զ����ļ�¼һ�������
        
        '���š���ˮ�š���š����롢���ơ���Ŀ���롢��񡢼�����λ�����������ۡ����Էѽ��
        gcn����.Execute "Insert Into " & mstrSavePath & "\YZ" & str��ˮ�� & " values ('','" & str��ˮ�� & "'," & _
            lng��� & ",'" & Trim(rsTemp!��Ŀ����) & "','" & Trim(rsTemp!����) & "','" & Trim(rsTemp!���) & "','" & Trim(rsTemp!���㵥λ) & "'," & _
            "''," & rs��ϸ!���� * rs��ϸ!���� & "," & rs��ϸ!ʵ�ս�� / (rs��ϸ!���� * rs��ϸ!����) & "," & rs��ϸ!ʵ�ս�� & "," & _
            "0)"
        lng��� = lng��� + 1
        gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rs��ϸ!ID & "')"
        Call ExecuteProcedure(gstrSysName)
        rs��ϸ.MoveNext
    Wend
    ���ʴ���_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

