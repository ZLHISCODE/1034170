Attribute VB_Name = "mdl�Ͼ���"
Option Explicit
Private mstrPatID As String
Private mobjSystem As New FileSystemObject
Private mobjStream As TextStream
Private mcur������� As Currency
Public gstr��ȷ���� As String

Private Type patInfo_�Ͼ���
    ����ʱ�� As String
    �������� As String
    ҽ������ As String
    ҽ������ As String
    ���ֱ��� As String
    �������� As String
    ҽ����������� As String
    ҽ����������� As String
    �����˱��� As String
End Type
Public gPatInfo_�Ͼ��� As patInfo_�Ͼ���

Private Type detailFee_�Ͼ���
    סԺ��� As String
    �������� As String
    ��־ As String
    ���÷���ʱ�� As String
    ҽԺ���� As String
    ҽԺ�Ա���  As String
    ҽ������ As String
    ���� As String
    ������λ As String
    ���� As Double
    ���� As Double
    �����˱��� As String
    ���� As String
    �������� As String
    ��� As String
End Type
Private mDetailFee_�Ͼ��� As detailFee_�Ͼ���

Private Type feeBalance_�Ͼ���
    סԺ��� As String
    ҽ������ As String
    ���÷���ʱ�� As String
    ������úϼ� As Double
    ҩ�Ѻϼ� As Double
    ������Ŀ�ϼ� As Double
    ������� As Double
    ҽ����Χ���� As Double
    �����ʻ�֧�� As Double
    ͳ��֧�� As Double
    ��֧�� As Double
    �����Ը� As Double
    �ڳ������ʻ� As Double
    ��ĩ�����ʻ� As Double
    ����Ա���� As String
    ���ݺ� As String
End Type
Public mFeeBalance As feeBalance_�Ͼ���

Public Function ҽ����ʼ��_�Ͼ���() As Boolean
     ҽ����ʼ��_�Ͼ��� = True
End Function

Public Function ��ݱ�ʶ_�Ͼ���(Optional bytType As Byte, Optional lng����id As Long) As String
    
    On Error GoTo errorhandle
    If bytType = 0 Then
        ��ݱ�ʶ_�Ͼ��� = frmIdentify�Ͼ���.Identify(bytType)
    Else
        ��ݱ�ʶ_�Ͼ��� = frm���ݽ���.getFeeBalance(bytType)
        Unload frm���ݽ���
    End If
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����������_�Ͼ���(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '�ֶΣ�������,����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��

    Dim rsTemp As New ADODB.Recordset, curCount As Currency
    Dim strFile As String, strWrite As String
    Dim strTemp As String
    
    'ɾ�����ܴ��ڵ�ǰ�ν�����Ϣ�ļ�
    On Error Resume Next
    Call Kill("C:\NJYB\MZJSHZ.TXT")
    
    On Error GoTo errorhandle
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�в��˷��ü�¼�����ܽ��н���", vbInformation, gstrSysName
        Exit Function
    End If
    curCount = 0
    While Not rs��ϸ.EOF
        curCount = curCount + rs��ϸ!ʵ�ս��
        rs��ϸ.MoveNext
    Wend
    rs��ϸ.MoveFirst
    
    'ȡ��������Ϣ��������
    mstrPatID = rs��ϸ!����ID
    With gPatInfo_�Ͼ���
        .����ʱ�� = Format(zlDatabase.Currentdate, "yyyyMMddHHmmss")        '�õ�����ʱ��
        .ҽ������ = Nvl(rs��ϸ!������)                                               '�õ�ҽ������
    End With
    
    If Trim(gPatInfo_�Ͼ���.ҽ������) = "" Then
        MsgBox "ҽ�������շѱ�������ҽ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "select A.�����ʼ� as ҽ������,C.���� as ҽ�����ұ���,C.���� as ҽ���������� from ��Ա�� A,������Ա B,���ű� C,�ٴ����� D " & _
              "where A.id=B.��Աid and B.����id = C.id and C.id=D.����id and B.ȱʡ=1 and  A.����='" & rs��ϸ!������ & "'"
    Call OpenRecordset(rsTemp, "ҽ������")
    If rsTemp.EOF Then
        MsgBox "δ��Ӧ���ҵ����ƿ�Ŀ����,������ȷ��Ӧ", vbInformation, gstrSysName
    End If
    
    With gPatInfo_�Ͼ���
        .ҽ������ = rsTemp!ҽ������                                               'ȡ��ҽ������
        .ҽ����������� = rsTemp!ҽ�����ұ���
        .ҽ����������� = rsTemp!ҽ����������
        .�����˱��� = UserInfo.���
    End With
    'д��ҽ��������Ϣ�ļ�
    strFile = "C:\NJYB\MZJZXX.TXT"
    strWrite = gPatInfo_�Ͼ���.����ʱ�� & fillSpa(gPatInfo_�Ͼ���.��������, 12) & _
               fillSpa(gPatInfo_�Ͼ���.ҽ������, 10) & fillSpa(gPatInfo_�Ͼ���.ҽ������, 8) & _
               fillSpa(gPatInfo_�Ͼ���.���ֱ���, 4) & fillSpa(gPatInfo_�Ͼ���.��������, 40) & _
               fillSpa(gPatInfo_�Ͼ���.ҽ�����������, 4) & fillSpa(gPatInfo_�Ͼ���.ҽ�����������, 30) & _
               fillSpa(gPatInfo_�Ͼ���.�����˱���, 10)
    Call writeTxtFile(strFile, strWrite)
    
    'ȡ����ϸ������������
    gstrSQL = "select ҽԺ���� from ������� where ���=" & TYPE_�Ͼ���
    Call OpenRecordset(rsTemp, "ҽԺ����")
    If rsTemp.EOF Then
        MsgBox "ҽԺ����δ����,��������ҽԺ����", vbInformation, gstrSysName
        Exit Function
    End If
    With mDetailFee_�Ͼ���
        .�������� = gPatInfo_�Ͼ���.��������
        .���÷���ʱ�� = gPatInfo_�Ͼ���.����ʱ��
        .ҽԺ���� = rsTemp!ҽԺ����
        .�����˱��� = gPatInfo_�Ͼ���.�����˱���
    End With
    
    '�ж��Ƿ���ҽ������δ��Ӧ
    Do Until rs��ϸ.EOF
        gstrSQL = "select A.��Ŀ����,B.���� from (select * from ����֧����Ŀ where ����=" & TYPE_�Ͼ��� & ") A, �շ�ϸĿ B where A.�շ�ϸĿid(+)=B.id and B.id = " & rs��ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, "ҽ����Ŀ")
        If IsNull((rsTemp!��Ŀ����)) Then
            MsgBox "<" & rsTemp!���� & ">δ��Ӧҽ������,���Ƚ��ж���", vbInformation, gstrSysName
            Exit Function
        End If
        rs��ϸ.MoveNext
    Loop
    
    strFile = "C:\NJYB\MZCFSJ.TXT"
    Call writeTxtFile(strFile, "")
    rs��ϸ.MoveFirst
    Do Until rs��ϸ.EOF
        gstrSQL = "select decode(A.���,'5',0,'6',0,'7',0,1) ��־,A.����,C.��Ŀ����,A.���㵥λ,B.����,decode(B.ҩƷ��Դ,'����',1,'����',2,'����',3,null) ��������,B.���" & _
                  " from �շ�ϸĿ A,ҩƷĿ¼ B,����֧����Ŀ C where A.id = C.�շ�ϸĿid and A.id=B.ҩƷid(+) and A.id =" & rs��ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, "��ϸ��ϸ")
        With mDetailFee_�Ͼ���
            .��־ = rsTemp!��־
            .���� = rsTemp!����
            .ҽ������ = rsTemp!��Ŀ����
            .������λ = zlCommFun.Nvl(rsTemp!���㵥λ)
            .���� = rs��ϸ!����
            .���� = rs��ϸ!����
            .���� = Nvl(rsTemp!����)
            .�������� = Nvl(rsTemp!��������)
            .��� = Nvl(rsTemp!���)
        End With
        strWrite = fillSpa(mDetailFee_�Ͼ���.��������, 12) & mDetailFee_�Ͼ���.��־ & _
                 mDetailFee_�Ͼ���.���÷���ʱ�� & _
                 fillSpa(mDetailFee_�Ͼ���.ҽ������, 40) & fillSpa(mDetailFee_�Ͼ���.����, 40) & _
                 fillSpa(mDetailFee_�Ͼ���.������λ, 10) & Lpad(mDetailFee_�Ͼ���.����, 10) & _
                 Lpad(Format(mDetailFee_�Ͼ���.����, "#0.00"), 10) & fillSpa(mDetailFee_�Ͼ���.�����˱���, 10) & _
                 fillSpa(mDetailFee_�Ͼ���.����, 20) & fillSpa(mDetailFee_�Ͼ���.��������, 1) & fillSpa(mDetailFee_�Ͼ���.���, 40)
        Call writeTxtFile(strFile, strWrite, False)
        rs��ϸ.MoveNext
    Loop
    Call writeTxtFile(strFile, "", False)
    
    '����ҽ��������
    strTemp = frm���ݽ���.getFeeBalance
    On Error Resume Next
    Unload frm���ݽ���
    On Error GoTo errorhandle
    If strTemp = "" Then
        MsgBox "��ȡҽ�������ļ����̱���ֹ,�޷����Ԥ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ȡ����ϢΪ���������׼��
    With mFeeBalance
        .ҽ������ = Val(analyseStr(strTemp, 1, 20))
        .������úϼ� = Val(analyseStr(strTemp, 35, 12))
        .������� = Val(analyseStr(strTemp, 67, 10))
        .ҽ����Χ���� = Val(analyseStr(strTemp, 77, 10))
        .�����ʻ�֧�� = Val(analyseStr(strTemp, 87, 10))
        .ͳ��֧�� = Val(analyseStr(strTemp, 97, 10))
        .��֧�� = Val(analyseStr(strTemp, 107, 10))
        .�����Ը� = Val(analyseStr(strTemp, 117, 10))
        .���ݺ� = Val(analyseStr(strTemp, 147, 20))
    End With
    If curCount <> CCur(mFeeBalance.������úϼ�) Then
        MsgBox "��ע�⣺ҽ�����ط��úϼ���ҽԺ������úϼƲ���" & vbCrLf & _
            "ҽԺ��" & curCount & Space(10) & "ҽ����" & mFeeBalance.������úϼ�
    End If
    mcur������� = Val(analyseStr(strTemp, 127, 10))
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mstrPatID & "," & TYPE_�Ͼ��� & ",'�ʻ����','" & mcur������� & "')"
    Call ExecuteProcedure(gstrSysName)
    
    str���㷽ʽ = "�����ʻ�;" & mFeeBalance.�����ʻ�֧�� & ";0"
    If mFeeBalance.ͳ��֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|ͳ�����;" & mFeeBalance.ͳ��֧�� & ";0"
    End If
    If mFeeBalance.��֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|��ͳ��;" & mFeeBalance.��֧�� & ";0"
    End If
'    If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 2)
    �����������_�Ͼ��� = True
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_�Ͼ���(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errorhandle
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�Ͼ��� & "," & mstrPatID & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              mFeeBalance.������úϼ� & "," & mFeeBalance.������� + mFeeBalance.�����Ը� & ",0," & _
              mFeeBalance.ҽ����Χ���� & "," & mFeeBalance.ͳ��֧�� & "," & mFeeBalance.��֧�� & "," & _
              "0," & mFeeBalance.�����ʻ�֧�� & ",null,null,null," & mFeeBalance.���ݺ� & ")"
    Call ExecuteProcedure("�Ͼ���ҽ��")
    
    �������_�Ͼ��� = True
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ����������_�Ͼ���(lng����ID As Long, cur�����ʻ� As Currency, lng����id As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long
    
    On Error GoTo errorhandle
    gstrSQL = "select distinct A.����id  from ���˷��ü�¼ A,���˷��ü�¼ B where A.��¼״̬=2 and A.NO=B.NO and B.����id=" & lng����ID
    Call OpenRecordset(rsTemp, "����id")
    lng����ID = rsTemp!����ID
    
    gstrSQL = "select * from ���ս����¼ where ��¼id=" & lng����ID
    Call OpenRecordset(rsTemp, "ԭʼ��¼")
    If rsTemp.EOF Then
        MsgBox "���ս����¼��ԭʼ���ʵ��ݲ�����,�������˷�", vbInformation, gstrSysName
        Exit Function
    Else
        gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�Ͼ��� & "," & rsTemp!����ID & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              -rsTemp!�������ý�� & "," & -rsTemp!ȫ�Ը���� & "," & -rsTemp!�����Ը���� & "," & -rsTemp!����ͳ���� & "," & -rsTemp!ͳ�ﱨ����� & "," & -rsTemp!���Ը���� & "," & _
              "0," & -rsTemp!�����ʻ�֧�� & ",null,null,null,null)"
        Call ExecuteProcedure("���ʼ�¼")
    End If
    
    ����������_�Ͼ��� = True
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_�Ͼ���(rsExse As Recordset, ByVal lng����id As Long) As String
    Dim bytType As Byte
    Dim strFile As String, strWrite As String
    Dim strStream As String
    Dim dblSettleSum As Double
    Dim rsTemp As New ADODB.Recordset
    'ɾ�����ܴ��ڵ�ǰ�ν�����Ϣ�ļ�
    On Error Resume Next
    Call Kill("C:\NJYB\CYJSD.TXT")
    On Error GoTo errorhandle
    '�ϴ���δ�ϴ�����ϸ����
    gstrSQL = "select ˳��� from �����ʻ� where ����id=" & lng����id
    Call OpenRecordset(rsTemp, "˳���")
    mDetailFee_�Ͼ���.סԺ��� = rsTemp!˳���
    
    '���ļ�
    strFile = "C:\NJYB\ZYFYMX.TXT"
    Call writeTxtFile(strFile, "")
    Do Until rsExse.EOF
        If rsExse!�Ƿ��ϴ� = 1 Then GoTo haddeliver             '�ҳ����ϴ���¼
        gstrSQL = "select decode(A.���,'5',0,'6',0,'7',0,1) ��־,A.����,A.����,C.��Ŀ����,A.���㵥λ,B.����,decode(B.ҩƷ��Դ,'����',1,'����',2,'����',3,null) ��������,B.���" & _
                  " from �շ�ϸĿ A,ҩƷĿ¼ B,����֧����Ŀ C where A.id = C.�շ�ϸĿid and A.id=B.ҩƷid(+) and A.id =" & rsExse!�շ�ϸĿID
        Call OpenRecordset(rsTemp, "��ϸ��ϸ")
        
        With mDetailFee_�Ͼ���
            .��־ = rsTemp!��־
            .���÷���ʱ�� = Format(rsExse!����ʱ��, "yyyyMMddHHmmss")
            .ҽԺ�Ա��� = rsTemp!����
            .ҽ������ = rsTemp!��Ŀ����
            .���� = rsTemp!����
            .������λ = zlCommFun.Nvl(rsTemp!���㵥λ)
            .���� = rsExse!�۸�
            .���� = rsExse!����
            .���� = zlCommFun.Nvl(rsTemp!����)
            .�������� = zlCommFun.Nvl(rsTemp!��������)
            .��� = zlCommFun.Nvl(rsTemp!���)
        End With
        
        gstrSQL = "select ����Ա��� from ���˷��ü�¼ where NO='" & rsExse!NO & "' and ���=" & rsExse!��� & _
                " and ��¼����=" & rsExse!��¼���� & " and ��¼״̬=" & rsExse!��¼״̬
        Call OpenRecordset(rsTemp, "����Ա���")
        mDetailFee_�Ͼ���.�����˱��� = rsTemp!����Ա���
        
        strWrite = mDetailFee_�Ͼ���.��־ & fillSpa(mDetailFee_�Ͼ���.סԺ���, 20) & _
                   mDetailFee_�Ͼ���.���÷���ʱ�� & _
                   fillSpa(mDetailFee_�Ͼ���.ҽ������, 40) & fillSpa(mDetailFee_�Ͼ���.����, 40) & _
                   fillSpa(mDetailFee_�Ͼ���.������λ, 10) & Lpad(mDetailFee_�Ͼ���.����, 10) & _
                   Lpad(Format(mDetailFee_�Ͼ���.����, "#0.00"), 10) & fillSpa(mDetailFee_�Ͼ���.�����˱���, 10) & _
                   fillSpa(mDetailFee_�Ͼ���.����, 20) & fillSpa(mDetailFee_�Ͼ���.��������, 1) & fillSpa(mDetailFee_�Ͼ���.���, 40)
        Call writeTxtFile(strFile, strWrite, False)
haddeliver:
        dblSettleSum = dblSettleSum + rsExse!���           '�ó������ܽ��
        rsExse.MoveNext
    Loop
    '�ر��ļ�
    Call writeTxtFile(strFile, "", False)
    
    bytType = 9                          '��ʾסԺԤ����״̬
    
    strStream = frm���ݽ���.getFeeBalance(bytType)
    On Error Resume Next
    Unload frm���ݽ���
    On Error GoTo errorhandle
    If strStream = "" Then
        MsgBox "��ȡҽ�������ļ����̱���ֹ,�޷����Ԥ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    With mFeeBalance
        .סԺ��� = analyseStr(strStream, 1, 20)
        .������úϼ� = Val(analyseStr(strStream, 35, 10))
        .ҽ����Χ���� = Val(analyseStr(strStream, 65, 10))
        .������� = Val(analyseStr(strStream, 75, 10))
        .�����Ը� = Val(analyseStr(strStream, 85, 10))
        .ͳ��֧�� = Val(analyseStr(strStream, 95, 10))
        .��֧�� = Val(analyseStr(strStream, 105, 10))
        .�����ʻ�֧�� = Val(analyseStr(strStream, 115, 10))
    End With
    
    If mFeeBalance.סԺ��� <> mDetailFee_�Ͼ���.סԺ��� Then
        MsgBox "�˽��ʲ�����ҽ�������ļ��в��˲�һ��,���ܽ���", vbInformation, gstrSysName
        Exit Function
    End If
    If Format(dblSettleSum, "#0.00") <> Format(mFeeBalance.������úϼ�, "#0.00") Then
        MsgBox "��ע��:ҽԺ�ܷ�����ҽ�����ķ��ص��ܷ��ò�һ��" & vbCrLf & _
        "�ܷ���:(ҽԺ)��" & Format(dblSettleSum, "#0.00") & Space(10) & "(ҽ��)��" & Format(mFeeBalance.������úϼ�, "#0.00"), vbInformation, gstrSysName
    End If

    strStream = "ͳ�����;" & mFeeBalance.ͳ��֧�� & ";0"
    If mFeeBalance.�����ʻ�֧�� <> 0 Then
        strStream = strStream & "|�����ʻ�;" & mFeeBalance.�����ʻ�֧�� & ";0"
    End If
    If mFeeBalance.��֧�� <> 0 Then
        strStream = strStream & "|��ͳ��;" & mFeeBalance.��֧�� & ";0"
    End If
    
    סԺ�������_�Ͼ��� = strStream
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_�Ͼ���(lng����ID As Long, lng����id) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errorhandle
    gstrSQL = "select NO,���,��¼״̬,��¼���� from ���˷��ü�¼ where nvl(�Ƿ��ϴ�,0)=0 and ����id=" & lng����ID
    Call OpenRecordset(rsTemp, "���Ҽ�¼")
    Do Until rsTemp.EOF
        gstrSQL = "ZL_���˷��ü�¼_�ϴ�('" & rsTemp!NO & "'," & rsTemp!��� & "," & rsTemp!��¼���� & "," & rsTemp!��¼״̬ & ")"
        Call ExecuteProcedure("�����ϴ���־")
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "select סԺ���� from ������Ϣ where ����id=" & lng����id
    Call OpenRecordset(rsTemp, "��ҳid")
    
    
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�Ͼ��� & "," & lng����id & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              mFeeBalance.������úϼ� & "," & mFeeBalance.������� + mFeeBalance.�����Ը� & ",0," & _
              mFeeBalance.ҽ����Χ���� & "," & mFeeBalance.ͳ��֧�� & "," & mFeeBalance.��֧�� & "," & _
              "0," & mFeeBalance.�����ʻ�֧�� & ",'" & mFeeBalance.סԺ��� & "'," & rsTemp!סԺ���� & ",null,null)"
    ExecuteProcedure ("���뱣���ʻ�")
    
    סԺ����_�Ͼ��� = True
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_�Ͼ���(lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long
    
    On Error GoTo errorhandle
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B where A.NO=B.NO and  A.��¼״̬=2 and B.ID=" & lng����ID
    Call OpenRecordset(rsTemp, "����id")
    lng����ID = rsTemp!ID
    
    gstrSQL = "select * from ���ս����¼ where ��¼id=" & lng����ID
    Call OpenRecordset(rsTemp, "ԭʼ��¼")
    If rsTemp.EOF Then
        MsgBox "���ս����¼��ԭʼ���ʵ��ݲ�����,�������˷�", vbInformation, gstrSysName
        Exit Function
    Else
        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�Ͼ��� & "," & rsTemp!����ID & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              -rsTemp!�������ý�� & "," & -rsTemp!ȫ�Ը���� & "," & -rsTemp!�����Ը���� & "," & -rsTemp!����ͳ���� & "," & -rsTemp!ͳ�ﱨ����� & "," & -rsTemp!���Ը���� & "," & _
              "0," & -rsTemp!�����ʻ�֧�� & ",'" & rsTemp!֧��˳��� & "'," & rsTemp!��ҳID & ",null,null)"
        ExecuteProcedure ("���ʼ�¼")
    End If
    
    סԺ�������_�Ͼ��� = True
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub writeTxtFile(strFile As String, strWrite As String, Optional ByVal openFile As Boolean = True)
    Dim intSymbol As Long
    Dim strFolder As String
    
    On Error GoTo errorhandle
    Do Until InStr(intSymbol + 1, strFile, "\") = 0
        intSymbol = InStr(intSymbol + 1, strFile, "\")
        strFolder = Mid(strFile, 1, intSymbol)
        If Not mobjSystem.FolderExists(strFolder) Then mobjSystem.CreateFolder (strFolder)
    Loop

    If openFile Then                    '���ļ�
        If Not mobjSystem.FileExists(strFile) Then mobjSystem.CreateTextFile (strFile)
        Set mobjStream = mobjSystem.OpenTextFile(strFile, ForWriting)
        If strWrite <> "" Then          '��������ݽ���д��
            mobjStream.WriteLine (UCase(strWrite))
            mobjStream.Close
        End If
    Else
        If strWrite = "" Then
            mobjStream.Close
        Else
            mobjStream.WriteLine (UCase(strWrite))   '�����д�����ݵ��򿪱�־Ϊfalse,ֻ����д��
        End If
    End If
    Exit Sub
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mobjStream.Close
End Sub

Public Function readTxtFile(strFile As String) As String
    On Error GoTo errHandle
    
    If mobjSystem.FileExists(strFile) Then
        Set mobjStream = mobjSystem.OpenTextFile(strFile)
        readTxtFile = mobjStream.ReadLine
        mobjStream.Close
    End If
    Exit Function
    
errHandle:
    Err.Clear
    On Error Resume Next
    mobjStream.Close
End Function

Private Function fillSpa(strTemp As Variant, lngLen As Long, Optional fromRigth As Boolean = True) As String
    Dim lngStrLeng As Long
    Dim strStream As String
    Dim strUnion As String
    
    strTemp = IIf(IsNull(strTemp), "", Trim(strTemp))
    
    strUnion = StrConv(Trim(strTemp), vbFromUnicode)
    lngStrLeng = IIf(LenB(strUnion) > lngLen, lngLen, LenB(strUnion))
    strStream = IIf(LenB(strUnion) > lngLen, StrConv(LeftB(strUnion, 20), vbUnicode), strTemp)
    
    If fromRigth Then
        fillSpa = strStream & String(lngLen - lngStrLeng, " ")
    Else
        fillSpa = String(lngLen - lngStrLeng, " ") & strStream
    End If
End Function

Public Function analyseStr(strTemp As String, lngStart As Long, lngLen As Long) As String
    Dim strStream As String
    
    strStream = StrConv(UCase(strTemp), vbFromUnicode)
    
    analyseStr = Trim(StrConv(MidB(strStream, lngStart, lngLen), vbUnicode))
End Function

Public Function �������_�Ͼ���(ByVal lng����id As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
'    gstrSQL = "select nvl(�ʻ����,0) as �ʻ���� from �����ʻ� where ����ID='" & lng����ID & "' and ����=" & TYPE_�Ͼ���
'    Call OpenRecordset(rsTemp, gstrSysName)
'
'    If rsTemp.EOF Then
'        �������_�Ͼ��� = 100000
'    Else
'        �������_�Ͼ��� = IIf(rsTemp("�ʻ����") = 0, 100000, rsTemp("�ʻ����"))
'    End If
    �������_�Ͼ��� = 100000
End Function

Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function GetFullNO(strNO As String) As String
'���ܣ����û�����Ĳ��ݵ��ţ����ص���ĵ��š�
    If Len(strNO) >= 8 Then GetFullNO = Right(strNO, 8): Exit Function
    GetFullNO = PreFixNO & Format(strNO, "0000000")
End Function

Public Function FileExists(ByVal FileName As String, Optional ErrFlag As Boolean = True) As Boolean
    Dim Temp
    FileExists = True
    On Error Resume Next
proshow:
    Temp = FileDateTime(FileName)
    Select Case Err
        Case 53, 76, 68
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                If ErrFlag Then
                    If MsgBox("����û��׼���á�", vbInformation + vbRetryCancel, "����") = vbRetry Then
                        GoTo proshow:
                    End If
                End If
                FileExists = False
            End If
    End Select
End Function
