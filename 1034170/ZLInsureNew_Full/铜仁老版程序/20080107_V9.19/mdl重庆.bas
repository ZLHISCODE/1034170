Attribute VB_Name = "mdl����"
Option Explicit
'API��������

'1���ӿڳ�ʼ��������������绷���Ƿ�ͨ������ҽԺ�ͻ�����ǰ�û���ǰ�û������ķ������䡣
Private Declare Function dy_Init Lib "SiInterface" Alias "INIT" () As Long

'2 ҵ��������ִ��ҽ��ҵ������Ҫ�Ĵ���
Private Declare Function dy_Business_Handle Lib "SiInterface" Alias "BUSINESS_HANDLE" _
    (ByVal InputData As String, ByVal OutputData As String) As Long
    
Private mstrҽ���� As String
Private mdbl��� As Double
Private mlng����id As Long
Private mstr����� As String
Private mblnIint As Boolean
Private mcnYB As New ADODB.Connection

Public Function ҽ����ʼ��_����() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    Dim lngReturn As Long
    
    If mblnIint = True Then
        'ֻ��Ҫ����һ��
        ҽ����ʼ��_���� = True
        Exit Function
    End If
    
    On Error Resume Next
    
    lngReturn = dy_Init
    If Err <> 0 Then
        MsgBox "������ȷ����ҽ���ӿڳ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    If lngReturn = -1 Then
        MsgBox "�������ҽ���ӿڳ�ʼ�������������������绷���Ƿ�ͨ��������" & vbCrLf & vbCrLf & _
          "1��ҽԺ�ͻ�����ҽԺǰ�û�Ӧ�÷�����֮�䣻" & vbCrLf & _
          "2��ҽԺǰ�û�Ӧ�÷�������ҽ������Ӧ�÷�����֮�䡣", vbInformation, gstrSysName
    Else
        ҽ����ʼ��_���� = True
        mblnIint = True
    End If
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    'alter table �����ʻ� add (������Ϣ varchar2(200),�������� varchar2(250));
    Dim strҽ���� As String, strInput As String, arrOutput  As Variant, int��� As Integer
    Dim str���� As String, str�Ա� As String, str���֤���� As String, lng���� As Long
    Dim str�������� As String, str��Ա��� As String, str��λ���� As String, str��λ���� As String
    Dim strIdentify As String, str���� As String, str���ı�� As String, str����� As String, str������Ϣ As String
    Dim datCurr As Date
    
    '��ʼ��һЩ����
    mlng����id = 0
    mstr����� = ""
    mstrҽ���� = ""
    mdbl��� = 0
    
    '����ǹҺţ���ֱ�ӵ������ﴦ��
    If bytType = 3 Then bytType = 0
    int��� = bytType
    If frmIdentify����.GetIdentify(strҽ����, int���) = False Then
        Exit Function
    End If
    
    '���ýӿ�
    strInput = "01|" & strҽ����
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    'ȡ�÷���ֵ
    str���� = arrOutput(1)
    str�Ա� = arrOutput(2)
    lng���� = Val(arrOutput(3))
    str���֤���� = arrOutput(4)
    str�������� = Get��������(str���֤����, lng����)
    
    str��Ա��� = ToVarchar(arrOutput(7), 8) 'VARCHAR2 (20) ��ְ����ְפ�⣬��ʱ�ù�������ְҵ��ת�ɣ����ݣ�������ؾ�ס����ְ����ְ��ؾ�ס��
    'arrOutput(8)   ����Ա��־               'VARCHAR2 (3)
    str��λ���� = ""
    str��λ���� = ToVarchar(arrOutput(9), 48) '50�ĳ��ȣ���Ҫ�۳�2������
    str���ı�� = arrOutput(10)
    
    If arrOutput(11) = "2" Then
        MsgBox "�ò���ҽ�������ܼ���ʹ�á�" & arrOutput(12)
        Exit Function
    End If
    
    str������Ϣ = ""
    If arrOutput(11) <> "0" Then
        'סԺʱҪ����
        str������Ϣ = arrOutput(12)
        MsgBox str������Ϣ, vbInformation, gstrSysName
    End If

    '����;ҽ����;����;����;�Ա�;��������;���֤;������λ
    'ҽ���ŵ�һλΪ������
    strIdentify = strҽ���� & ";" & strҽ���� & ";;" & str���� & ";" & str�Ա� & ";" & str�������� & ";" & str���֤���� & ";" & str��λ���� & "(" & str��λ���� & ")"
    strIdentify = Replace(strIdentify, " ", "")
    
    str���� = ";"                                       '8.���Ĵ���
    str���� = str���� & ";"                             '9.˳���
    str���� = str���� & ";" & str��Ա���               '10��Ա���
    str���� = str���� & ";0"                            '11�ʻ����
    str���� = str���� & ";0"                            '12��ǰ״̬
    str���� = str���� & ";"                             '13����ID
    str���� = str���� & ";" & IIf(Left(str��Ա���, 1) = "��", 2, 1)     '14��ְ(1,2)
    str���� = str���� & ";"                             '15����֤��
    str���� = str���� & ";" & lng����                   '16�����
    str���� = str���� & ";"                             '17�Ҷȼ�
    str���� = str���� & ";0"                            '18�ʻ������ۼ�
    str���� = str���� & ";0"                            '19�ʻ�֧���ۼ�
    str���� = str���� & ";"                             '20����ͳ���ۼ�
    str���� = str���� & ";"                             '21ͳ�ﱨ���ۼ�
    str���� = str���� & ";"                             '22סԺ�����ۼ�
    str���� = str���� & ";" & IIf(int��� = 14, 1, "")  '23�������� (1����������)
    
    lng����ID = BuildPatiInfo(bytType, strIdentify & str����, lng����ID)
    
    mdbl��� = 0
    If bytType = 0 Then        '��������ͬʱ���о���Ǽ�
        '��������ⲡ�������ȣ���Ҫѡ���˼���
        Dim rs���� As ADODB.Recordset
        Dim str���� As String, str�������� As String, str����֢ As String, str�������� As String
        
        If int��� = 13 Or int��� = 14 Then
            If int��� = 13 Then
                '���������Ϣ
                strInput = "07|" & strҽ����
                If HandleBusiness(strInput, arrOutput) = False Then Exit Function
                
                str���� = "���ⲡ"
                If frm����ѡ������.GetCode(arrOutput, str����, str��������, str��������, str����֢) = False Then Exit Function
            Else
                str���� = "����"
                If frm����ѡ������.GetCode("", str����, str��������, str��������, str����֢) = False Then Exit Function
            End If
    
            gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'����֤��','''" & str�������� & "''')"
            Call ExecuteProcedure("���²���")
            gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'��������','''" & str�������� & "''')"
            Call ExecuteProcedure("���²���")
            gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'����֢','''" & str����֢ & "''')"
            Call ExecuteProcedure("���²���֢")
        End If
        
        datCurr = zlDatabase.Currentdate
        str����� = ToVarchar(lng����ID & Format(datCurr, "yyMMddHHmmss"), 18)
        strInput = "02|" & str����� & "|" & int��� & "|" & strҽ���� & _
                   "|����|" & ToVarchar(UserInfo.����, 20) & "|" & _
                   Format(datCurr, "yyyy-MM-dd") & "|" & str�������� & "|" & ToVarchar(UserInfo.����, 20) & "|" & str����֢
        If HandleBusiness(strInput, arrOutput) = False Then Exit Function
        
        mlng����id = lng����ID
        mstr����� = str�����
        mstrҽ���� = strҽ����
        mdbl��� = Val(arrOutput(2))
    End If
    
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(datCurr) & "," & _
        mdbl��� & ",0,0,0,0,0,0,0,0,0,'')"
    Call ExecuteProcedure("����ҽ��")
    
    '���·�����Ϣ
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'������Ϣ','''" & str������Ϣ & "''')"
    Call ExecuteProcedure("���²���")
    
    g��������.�����Ը���� = int��� '������ʱ���棬�������
    
    '���ظ�ʽ:�м���벡��ID
    If lng����ID <> 0 Then
        ��ݱ�ʶ_���� = strIdentify & ";" & lng����ID & str����
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(strSelfNo As String) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: strSelfNO-���˸��˱��
'����: ���ظ����ʻ����Ľ��
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHandle
    
    '�����ݿ��ж�ȡ����Ϊ�ղŲű����˵ģ�Ӧ����׼ȷ�ģ�
    If mstrҽ���� = "" Or strSelfNo <> mstrҽ���� Then
        gstrSQL = "Select �ʻ���� From �����ʻ� where ����=" & gintInsure & " and ����=0 and ҽ����='" & strSelfNo & "'"
        Call OpenRecordset(rsTemp, "����ҽ��")
        
        If rsTemp.EOF = False Then
            �������_���� = IIf(IsNull(rsTemp("�ʻ����")), 0, rsTemp("�ʻ����"))
        End If
    Else
        �������_���� = mdbl���
    End If
    'ֻ����һ��
    mstrҽ���� = ""
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
'������rsDetail     ������ϸ(����)
'      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Static str�����Pre As String
    Dim strҽ���� As String, strInput As String, arrOutput  As Variant
    Dim dbl�����ʻ� As Double, strMessage As String
    Dim lng����ID As Long, str��� As String, datCurr As Date
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If rs��ϸ.RecordCount = 0 Then
        str���㷽ʽ = "�����ʻ�;0;0"
        �����������_���� = True
        Exit Function
    End If
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ("����ID")
    datCurr = zlDatabase.Currentdate
    
    If mlng����id <> lng����ID Then
        MsgBox "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�����˵���ǰ����������δ��ķ��ã��������ִ��Ԥ����
    If str�����Pre = mstr����� Then
        '�Ѿ���ֵ��˵���ò��˽��й�Ԥ��
        strInput = "10|" & mstr����� & "|" & mstr�����
        If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    End If
    '�����ֵ
    str�����Pre = mstr�����
    
    'Ȼ����봦����ϸ
    Do Until rs��ϸ.EOF
        gstrSQL = "select A.����,A.����,A.���,A.���,A.���㵥λ,B.��Ŀ����,B.��ע,A.���㵥λ,E.���,G.���� ���� " & _
                  "from �շ�ϸĿ A,����֧����Ŀ B,ҩƷĿ¼ E ,ҩƷ��Ϣ F,ҩƷ���� G " & _
                  "where A.ID=" & rs��ϸ("�շ�ϸĿID") & " and A.ID=B.�շ�ϸĿID and B.����=" & gintInsure & _
                 "        AND A.ID=E.ҩƷID(+) AND E.ҩ��ID=F.ҩ��ID(+) AND F.����=G.����(+) "
        Call OpenRecordset(rsTemp, "����Ԥ��")
        If rsTemp.EOF = True Then
            MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
            Exit Function
        End If
        If Val(Nvl(rs��ϸ("ʵ�ս��"), 0)) <> 0 Then
            strInput = "04|" & mstr����� & "|" & mstr����� & "|" & Format(datCurr, "yyyy-MM-dd HH:mm:ss")
            strInput = strInput & "|" & ToVarchar(rsTemp("��Ŀ����"), 10)  'ҽ����ˮ��
            strInput = strInput & "|" & ToVarchar(rsTemp("����"), 20)      'ҽԺ����
            strInput = strInput & "|" & ToVarchar(rsTemp("����"), 50)      '��Ŀ����
            strInput = strInput & "|" & Format(rs��ϸ("����"), "0.0000")   '����
            strInput = strInput & "|" & Format(rs��ϸ("����"), "0.00")     '����
            strInput = strInput & "|" & IIf(rs��ϸ("�Ƿ���") = 1, 1, 0)  '�����־
            strInput = strInput & "|" & Format(UserInfo.����, 20)          '����ҽ��
            strInput = strInput & "|" & Format(UserInfo.����, 20)          '������
            strInput = strInput & "|" & ToVarchar(rsTemp("���㵥λ"), 20)     '��λ
            strInput = strInput & "|" & ToVarchar(rsTemp("���"), 14)         '���
            strInput = strInput & "|" & ToVarchar(rsTemp("����"), 20)         '����
            strInput = strInput & "|"                                         '������ϸ��ˮ��
            strInput = strInput & "|" & Format(rs��ϸ("ʵ�ս��"), "#####0.0000")         '���
            
            If HandleBusiness(strInput, arrOutput) = False Then Exit Function
            Call AddMessage(strMessage, arrOutput, ToVarchar(rsTemp("����"), 50), rs��ϸ("����"))
        End If
        rs��ϸ.MoveNext
    Loop
    
    If strMessage <> "" Then
        strMessage = "���˷�����ϸ��������еõ�ҽ���������·�����Ϣ���Ƿ������" & vbCrLf & vbCrLf & strMessage
        If MsgBox(strMessage, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            '�û�ѡ��ȡ�������˵���ϸ
            strInput = "10|" & mstr����� & "|" & mstr�����
            If HandleBusiness(strInput, arrOutput) = False Then Exit Function
                        
            Exit Function
        End If
    End If
    '����Ԥ����
    
    strInput = "06|" & mstr�����
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    
    str���㷽ʽ = "�����ʻ�;" & Val(arrOutput(2)) & ";0"  '�����޸ĸ����ʻ�����Ϊ����ʱ�Ѿ����ٴ���ǰ�û���
    If Val(arrOutput(1)) > 0 Then
        str���㷽ʽ = str���㷽ʽ & "|ҽ������;" & Val(arrOutput(1)) & ";0"
    End If
    If Val(arrOutput(3)) > 0 Then
        str���㷽ʽ = str���㷽ʽ & "|����Ա����;" & Val(arrOutput(3)) & ";0"
    End If
    If Val(arrOutput(5)) > 0 Then
        str���㷽ʽ = str���㷽ʽ & "|���ͳ��;" & Val(arrOutput(5)) & ";0"
    End If
    If Val(arrOutput(6)) > 0 Then
        str���㷽ʽ = str���㷽ʽ & "|����Ա����;" & Val(arrOutput(6)) & ";0"
    End If
    
    �����������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim strҽ���� As String, strInput As String
    Dim lng����ID  As Long
    Dim str����Ա As String, arrOutput  As Variant
    Dim cur��������, datCurr As Date
    Dim str���� As String, str����֢ As String, str������Ϣ As String, str��ע As String
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    Call DebugTool("�����������")
    gstrSQL = "Select * From ���˷��ü�¼ Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Call OpenRecordset(rs��ϸ, "����ҽ��")
    
    If rs��ϸ.EOF = True Then
        MsgBox "û����д�շѼ�¼", vbExclamation, gstrSysName
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")
    str����Ա = ToVarchar(IIf(IsNull(rs��ϸ("����Ա����")), UserInfo.����, rs��ϸ("����Ա����")), 20)
    
    If mlng����id <> lng����ID Then
        MsgBox "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    Do Until rs��ϸ.EOF
        cur�������� = cur�������� + rs��ϸ("���ʽ��")
        Call TrackRecordInsure(rs��ϸ!ID, rs��ϸ!�շ�ϸĿID)
        rs��ϸ.MoveNext
    Loop
    
    '��ȡ�α����˵Ĳ���֢��������Ϣ
    gstrSQL = "Select ����֤�� As ���ֱ���,��������,����֢,������Ϣ" & _
        " From �����ʻ� " & _
        " Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ�α����˵Ĳ���֢��������Ϣ")
    If Not rsTemp.EOF Then
        str���� = Nvl(rsTemp!���ֱ���)
        If str���� <> "" Then str���� = "[" & str���� & "]"
        str���� = str���� & Nvl(rsTemp!��������)
        str����֢ = Nvl(rsTemp!����֢)
        str������Ϣ = Nvl(rsTemp!������Ϣ)
    End If
    str��ע = str���� & "||" & str����֢ & "||" & str������Ϣ
    Call DebugTool("�ɹ���ȡ���������Ϣ�������ֱ��������ơ�����֢��������Ϣ")
    
    '���ý���
    Call DebugTool("׼�������������")
    strInput = "05|" & mstr����� & "|1||" & str����Ա & "|0" '���ʻ����֧��
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    Call DebugTool("���ս����¼")
    '��������¼
    '---------------------------------------------------------------------------------------------
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
            
    Dim curͳ��֧�� As Double
    Dim cur����Ա���� As Double
    Dim cur���ͳ�� As Double
    
    cur�������� = Val(Format(cur��������, "#####0.00"))
    curͳ��֧�� = Val(arrOutput(2))
    cur�����ʻ� = Val(arrOutput(3))
    cur����Ա���� = Val(arrOutput(4))
    cur���ͳ�� = Val(arrOutput(6))
    
    '�ʻ������Ϣ
    datCurr = zlDatabase.Currentdate
    
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
                
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� + curͳ��֧�� & "," & _
        curͳ�ﱨ���ۼ� + curͳ��֧�� & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure("����ҽ��")
    
    'g��������.�����Ը�����б���������ﲡ�˾������ͣ�������ⲡ�������ͨ����������¼�ı�ע������ǲ��ֵ�����
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)'�����Ը����������ʱ���棬�������
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & cur�������� & ",0,0," & _
        curͳ��֧�� & "," & curͳ��֧�� & ",0," & g��������.�����Ը���� & "," & cur�����ʻ� & ",'" & arrOutput(1) & "',NULL,NULL,'" & str��ע & "')"
    Call ExecuteProcedure("����ҽ��")
    '---------------------------------------------------------------------------------------------

    �������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, strInput As String, arrOutput  As Variant
    Dim lng����ID As Long, str��ˮ�� As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim curƱ���ܽ�� As Currency
    Dim curDate As Date
        
    On Error GoTo errHandle
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "Select * From ���˷��ü�¼ " & _
        " Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Call OpenRecordset(rsTemp, "����ҽ��")
    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    curƱ���ܽ�� = Val(Format(curƱ���ܽ��, "#####0.00"))
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "����ҽ��")
    lng����ID = rsTemp("����ID")
    
    '����ҽ����Ŀ��״̬������
    gstrSQL = "Select * From ���˷��ü�¼ " & _
        " Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Call OpenRecordset(rsTemp, "����ҽ��")
    Do While Not rsTemp.EOF
        Call TrackRecordInsure(rsTemp!ID, rsTemp!�շ�ϸĿID)
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=" & gintInsure & " and ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, "����ҽ��")
    If rsTemp.EOF = True Then
        MsgBox "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    str��ˮ�� = rsTemp("֧��˳���")
    cur�����ʻ� = Nvl(rsTemp!�����ʻ�֧��, 0)
    
    strInput = "99|" & str��ˮ�� & "|" & ToVarchar(UserInfo.����, 20)
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure("����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0," & rsTemp("�����Ը����") & "," & _
        cur�����ʻ� * -1 & ",'" & str��ˮ�� & "')"
    Call ExecuteProcedure("����ҽ��")

    ����������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����ʻ�תԤ��_����(lngԤ��ID As Long, cur�����ʻ� As Currency, strSelfNo As String, str˳��� As String, ByVal lng����ID As Long) As Boolean
'���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
    
    �����ʻ�תԤ��_���� = False
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim strInput As String, arrOutput  As Variant
    Dim datCurr As Date, rsTemp As New ADODB.Recordset
    Dim str���� As String, str˳��� As String
    Dim strTemp As String, str��ʾ As String, str��� As String
    
    On Error GoTo errHandle
    
    '��ò��˳�Ժ���
    gstrSQL = "select A.������Ϣ from ������ A where A.����ID=" & lng����ID & " and A.��ҳID=" & lng��ҳID & _
              " and A.�������=1 and A.��ϴ���=1"
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    If rsTemp.EOF = False Then
        str��� = ToVarchar(rsTemp("������Ϣ"), 40)
    End If
    
    '���ҽ����
    gstrSQL = "select ����,ҽ���� from �����ʻ� where ����=" & TYPE_������ & " and ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    str���� = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
    strҽ���� = rsTemp("ҽ����")
    
    '���������Ժ��Ϣ
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select A.��Ժ��ʽ,nvl(A.����Ժת��,0) as ����Ժת��,A.����ҽʦ,A.��Ժ����,A.��Ժ����,B.���� as ��Ժ���� from ������ҳ A,���ű� B " & _
             " Where A.��Ժ����ID = B.ID And A.����ID =" & lng����ID & " And A.��ҳID = " & lng��ҳID
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    
    '������Ժ�ӿ�
    strInput = "02|" & GetIdentify(lng����ID, lng��ҳID) & "|" & IIf(rsTemp("��Ժ��ʽ") = "ת��", "22", "21") & "|" & strҽ���� & "|" & _
               ToVarchar(rsTemp("��Ժ����"), 30) & "|" & ToVarchar(rsTemp("����ҽʦ"), 20) & "|" & _
               Format(rsTemp("��Ժ����"), "yyyy-MM-dd") & "|" & ToVarchar(str���, 40) & "|" & ToVarchar(UserInfo.����, 20) & "|0"
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    str˳��� = arrOutput(1)
    mdbl��� = Val(arrOutput(2))
    
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        mdbl��� & ",0,0,0,0,0,0,0,0,0,'')"
    Call ExecuteProcedure("����ҽ��")
    
    'ǿ�ưѵǼ�˳��š����µ�ҽ��������
    gstrSQL = "ZL_�����ʻ�_�޸�ҽ����(" & lng����ID & "," & gintInsure & _
                ",'" & str���� & "','" & strҽ���� & "','" & str˳��� & "')"
    Call ExecuteProcedure("����ҽ��")
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure("����ҽ��")
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
            'ȡ��Ժ�Ǽ���֤�����ص�˳���
    Dim datCurr As Date, rsTemp As New ADODB.Recordset
    Dim strInput As String, arrOutput  As Variant, bln����ó�Ժ As Boolean
    Dim str��� As String
    Dim strסԺ����� As String
    
    On Error GoTo errHandle
    
    '��ò��˳�Ժ���
    gstrSQL = "select A.������Ϣ from ������ A where A.����ID=" & lng����ID & " and A.��ҳID=" & lng��ҳID & _
              " and A.�������=3 and A.��ϴ���=1"
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    If rsTemp.EOF = False Then
        str��� = Nvl(rsTemp("������Ϣ"), "��")
    Else
        str��� = "��"   '��ϲ�����β���Ϊ��
    End If
    str��� = ToVarchar(str���, 40)
    
    '���������Ժ��Ϣ
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select A.סԺҽʦ,A.��Ժ����,A.��Ժ����,A.��Ժ����,B.���� as ��Ժ���� from ������ҳ A,���ű� B " & _
             " Where A.��Ժ����ID = B.ID And A.����ID =" & lng����ID & " And A.��ҳID = " & lng��ҳID
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    '���ýӿڣ�������סԺ��Ϣ
    strInput = "03|" & GetIdentify(lng����ID, lng��ҳID) & "|0001010010|21|||" & Format(rsTemp("��Ժ����"), "yyyy-MM-dd") & "||" & _
                Format(rsTemp("��Ժ����"), "yyyy-MM-dd") & "|||" & ToVarchar(UserInfo.����, 20) & "|0"
    
    '���ô�סԺ�Ƿ�û�з��÷���
    gstrSQL = "Select nvl(sum(ʵ�ս��),0) as ���  from ���˷��ü�¼ where ����ID=" & lng����ID & " and ��ҳID=" & lng��ҳID
    Call OpenRecordset(rsTemp, "���˳�Ժ")
    If rsTemp.EOF = True Then
        bln����ó�Ժ = True
    Else
        bln����ó�Ժ = (rsTemp("���") = 0)
    End If
    
    If bln����ó�Ժ = True Then
        '��������ó�Ժ���ͽ��䴦��Ϊ����Ժ�������ø�����סԺ��Ϣ
        gstrSQL = "Select ˳��� from �����ʻ� where ����ID=" & lng����ID & " and ����=" & gintInsure
        Call OpenRecordset(rsTemp, "���˳�Ժ")
        strInput = "99|" & rsTemp("˳���") & "|" & ToVarchar(UserInfo.����, 20)
    End If
    
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure("����ҽ��")
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure("����ҽ��")
    
    ��Ժ�Ǽǳ���_���� = True
End Function

Public Function ���³�Ժ����_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ����²��˵ĳ�Ժ����������������������ʱ���߻����
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str���� As String, str����֢ As String, str�������� As String, str�������� As String
    Dim strInput As String, arrOutput  As Variant
    Dim str���� As String, strҽ�� As String
    Dim str��Ժ���� As String, str��Ժ���� As String, str��� As String, strҽ����� As String
    
    On Error GoTo errHandle
    
    '��ò��˳�Ժ���ּ�����֢
    gstrSQL = "Select ����֤�� ���ֱ���,��������,����֢ From �����ʻ� Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ���˳�Ժ���ּ�����֢")
    str�������� = Nvl(rsTemp!���ֱ���)
    str����֢ = Nvl(rsTemp!����֢)
    str�������� = Nvl(rsTemp!��������)
    
    str���� = "��Ժ"
    If frm����ѡ������.GetCode("", str����, str��������, str��������, str����֢) = False Then
        Exit Function
    End If
    str�������� = ToVarchar(str��������, 20)
    str����֢ = ToVarchar(str����֢, 200)
    str�������� = TrimStr(str��������)
    
    '��ò��˳�Ժ���
    gstrSQL = "select A.������Ϣ from ������ A where A.����ID=" & lng����ID & " and A.��ҳID=" & lng��ҳID & _
              " and A.�������=1 and A.��ϴ���=1"
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    If rsTemp.EOF = False Then
        str��� = ToVarchar(rsTemp("������Ϣ"), 40)
    End If
    
    'ȡ���˵���Ժ��Ϣ
    gstrSQL = "select A.��Ժ��ʽ,nvl(A.����Ժת��,0) as ����Ժת��,A.����ҽʦ,A.��Ժ����,A.��Ժ����,A.��Ժ����,B.���� as ��Ժ���� from ������ҳ A,���ű� B " & _
             " Where A.��Ժ����ID = B.ID And A.����ID =" & lng����ID & " And A.��ҳID = " & lng��ҳID
    Call OpenRecordset(rsTemp, "ȡ���˵���Ժ��Ϣ")
    strҽ����� = IIf(rsTemp!��Ժ��ʽ = "ת��", 22, 21)
    str���� = ToVarchar(rsTemp("��Ժ����"), 30)
    strҽ�� = ToVarchar(rsTemp("����ҽʦ"), 20)
    str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy-MM-dd")
    If IsNull(rsTemp!��Ժ����) Then
        str��Ժ���� = ""
    Else
        str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy-MM-dd")
    End If
    
    'Modified by ZYB 2004-05-10
    '���ýӿڸ��£�סԺ��(�����)|���±�־|ҽ�����|����|ҽ��|��Ժ����|��Ժ���|��Ժ����|ȷ�Ｒ������
                  '|��Ժԭ��|������|����֢
    strInput = "03|" & GetIdentify(lng����ID, lng��ҳID) & "|1111111001|" & strҽ����� & "|" & str���� & _
               "|" & strҽ�� & "|" & str��Ժ���� & "|" & str��� & "|" & str��Ժ���� & "|" & str�������� & "|||" & str����֢
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'����֤��','''" & str�������� & "''')"
    Call ExecuteProcedure("���²���")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'��������','''" & str�������� & "''')"
    Call ExecuteProcedure("���²���")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'����֢','''" & str����֢ & "''')"
    Call ExecuteProcedure("���²���֢")
    
    ���³�Ժ����_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ����ҽ����Ժ_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str˳��� As String) As Boolean
'���ܣ����²��˵ĳ�Ժ����������������������ʱ���߻����
    Dim strInput As String, arrOutput  As Variant
    
    On Error GoTo errHandle
    
    '���ýӿ�
    strInput = "99|" & str˳��� & "|" & ToVarchar(UserInfo.����, 20)
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
    Call ExecuteProcedure("����ҽ����Ժ")
    
    ����ҽ����Ժ_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����(rsExse As Recordset, ByVal lng����ID As Long, ByVal strҽ���� As String, _
        Optional ByVal bln���ò�ѯ As Boolean = False) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim cn�ϴ� As New ADODB.Connection, rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    
    Dim str������Ϣ As String
    Dim strInput As String, arrOutput   As Variant, arrTemp As Variant
    Dim cur�����ʻ� As Double, curͳ��֧�� As Double, cur���ͳ�� As Double, cur����Ա���� As Double, cur�������� As Double
    Dim str�ܽ��ҽԺ As String, str�ܽ��ҽ�� As String, str������ϸ��ˮ�� As String
    Dim strҽ�� As String, datCurr As Date, intMsg As Integer
    
    On Error GoTo errHandle
    mlng����id = 0         '��ʼ����ֻҪһѡ���ˣ��ͻ���ñ����̣�Ҳ�ͻ����0
    
    If rsExse.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If
    rsExse.MoveFirst
    
    datCurr = zlDatabase.Currentdate
    With g��������
        .����ID = rsExse("����ID")
        
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=" & rsExse("����ID")
        Call OpenRecordset(rsTemp, "�������")
        If IsNull(rsTemp("��ҳID")) = True Then
            MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
            Exit Function
        End If
        .��ҳID = rsTemp("��ҳID")
        .��� = Int(Format(datCurr, "yyyy"))
    End With
    
    'Modified by ZYB 2004-05-10
    '��ȡ���˵Ļ�����Ϣ��������ڷ���ԭ������ʾ
    gstrSQL = "Select ҽ���� From �����ʻ� Where ����=" & TYPE_������ & " And ����ID=" & g��������.����ID
    Call OpenRecordset(rsTemp, "ȡ�ò��˵�ҽ����")
    
    strInput = "01|" & rsTemp!ҽ����
    If HandleBusiness(strInput, arrTemp) = False Then Exit Function
    str������Ϣ = ""
    If Val(arrTemp(11)) <> 0 Then
        str������Ϣ = arrTemp(12)
        MsgBox str������Ϣ, vbInformation, gstrSysName
    End If
    '���·�����Ϣ
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'������Ϣ','''" & str������Ϣ & "''')"
    Call ExecuteProcedure("���²���")
    
    Screen.MousePointer = vbHourglass
    '1.2 �������˵���Ժʱ��
    gstrSQL = "select ��Ժ����,nvl(��Ժ����,to_date('3000-01-01','yyyy-MM-dd')) as ��Ժ���� " & _
              "from ������ҳ where ����ID=" & g��������.����ID & " and ��ҳID=" & g��������.��ҳID
    Call OpenRecordset(rsTemp, "�������")
    If rsTemp("��Ժ����") = CDate("3000-01-01") Then
        g��������.��;���� = 1
        g��������.סԺ���� = DateDiff("d", rsTemp("��Ժ����"), datCurr)
    Else
        '��ʾ�ò����Ѿ���Ժ
        g��������.��;���� = 0
        g��������.סԺ���� = DateDiff("d", rsTemp("��Ժ����"), rsTemp("��Ժ����"))
    End If
    If g��������.סԺ���� < 1 Then g��������.סԺ���� = 1 '������һ��
    
    Do Until rsExse.EOF
        cur�������� = cur�������� + rsExse("���")
        rsExse.MoveNext
    Loop
    cur�������� = Val(Format(cur��������, "#####0.00"))
    
    'ֻ�г�Ժ�����ʹ�ò��˷��ò�ѯ�����ã����ϴ�����δ�ϴ���ϸ����;����ֻ�����ϴ����ݽ��н���
    If g��������.��;���� = 0 Or bln���ò�ѯ Then
        '����δ�ϴ���ϸ�������Ա����ϴ�����ϸ�����ϴ�����ϸ��
        gstrSQL = "Select A.ID,A.NO,A.��¼����,A.��¼״̬,A.���,A.����ID,A.��ҳID,A.����ʱ�� as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս��" & _
                  "         ,A.�շ�ϸĿID,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
                  "         ,C.��Ŀ����,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,A.����Ա����,B.���㵥λ,E.���,G.���� ���� " & _
                  "  From ���˷��ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C,������ҳ D,ҩƷĿ¼ E ,ҩƷ��Ϣ F,ҩƷ���� G " & _
                  "  where A.����ID=" & lng����ID & " and A.��ҳID=" & g��������.��ҳID & " and A.���ʷ���=1 and A.ʵ�ս��<>0 and nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 " & _
                  "        and A.����ID=D.����ID and A.��ҳID=D.��ҳID And D.����=" & gintInsure & _
                  "        and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����=D.���� " & _
                  "        AND B.ID=E.ҩƷID(+) AND E.ҩ��ID=F.ҩ��ID(+) AND F.����=G.����(+) " & _
                  "  Order by A.����ʱ��,A.��¼����,Decode(A.��¼״̬,2,2,1)"
        Call OpenRecordset(rs��ϸ, "�������")
        
        '����ҽ����Ŀ����Ϣ������
        Do While Not rs��ϸ.EOF
            Call TrackRecordInsure(rs��ϸ!ID, rs��ϸ!�շ�ϸĿID)
            rs��ϸ.MoveNext
        Loop
        
        '������һ�����Ӵ����Դﵽ���ܵ�ǰ��������Ŀ���
        cn�ϴ�.ConnectionString = gcnOracle.ConnectionString
        cn�ϴ�.Open
        
        intMsg = 0
        Call DebugTool("��ʼ�ϴ���ϸ")
        If rs��ϸ.RecordCount <> 0 Then rs��ϸ.MoveFirst
        Do Until rs��ϸ.EOF
            If rs��ϸ!��¼״̬ = 1 Then
                If Val(rs��ϸ!����) < 0 Or Val(rs��ϸ!�۸�) < 0 Then
                    '����ȡһ��������¼����ˮ�ţ���Ϊ������ˮ��
                    str������ϸ��ˮ�� = GetSequence(rs��ϸ!����ID, rs��ϸ!��ҳID, rs��ϸ!�շ�ϸĿID)
                Else
                    str������ϸ��ˮ�� = ""
                End If
            Else
                str������ϸ��ˮ�� = GetDetailSequence(rs��ϸ!NO, rs��ϸ!���, rs��ϸ!��¼����, rs��ϸ!��¼״̬)
            End If
            strҽ�� = ToVarchar(IIf(IsNull(rs��ϸ("ҽ��")), UserInfo.����, rs��ϸ("ҽ��")), 20)
            
            strInput = "04|" & GetIdentify(lng����ID, g��������.��ҳID)
            strInput = strInput & "|" & rs��ϸ("NO") & "_" & rs��ϸ("��¼����")
            strInput = strInput & "|" & Format(rs��ϸ("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")
            strInput = strInput & "|" & ToVarchar(rs��ϸ("��Ŀ����"), 10) '���ı���
            strInput = strInput & "|" & ToVarchar(rs��ϸ("����"), 20) 'ҽԺ����
            strInput = strInput & "|" & ToVarchar(rs��ϸ("����"), 50)     '��Ŀ����
            strInput = strInput & "|" & Format(rs��ϸ("�۸�"), "0.0000")      '����
            strInput = strInput & "|" & Format(rs��ϸ("����"), "0.00")        '����
            strInput = strInput & "|" & IIf(rs��ϸ("�Ƿ���") = 1, 1, 0)     '�����־
            strInput = strInput & "|" & strҽ��                               'ҽ��
            strInput = strInput & "|" & ToVarchar(UserInfo.����, 20)          '������
            strInput = strInput & "|" & ToVarchar(rs��ϸ("���㵥λ"), 20)     '��λ
            strInput = strInput & "|" & ToVarchar(rs��ϸ("���"), 14)         '���
            strInput = strInput & "|" & ToVarchar(rs��ϸ("����"), 20)         '����
            strInput = strInput & "|" & str������ϸ��ˮ��                     '������ϸ��ˮ��
            strInput = strInput & "|" & Format(rs��ϸ("ʵ�ս��"), "#####0.0000")         '���
            
            'Modified by ZYB 20040511 ��������
            '�������ڸ������ʣ������ĳ�����¼����Ϊ���������Ǳ����ڽӿ����ƣ��϶�������ȥ����˱��������������ĳ�����¼���ϴ�
            If Not (rs��ϸ!��¼״̬ = 2 And str������ϸ��ˮ�� = "") Then
                If HandleBusiness(strInput, arrOutput) = False Then
                    '�����ϴ�ʧ��
                    If MsgBox("���ݡ�" & rs��ϸ("NO") & "����" & rs��ϸ("����") & "�����ϴ�ʧ�ܣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                    If intMsg = 0 Then
                        If MsgBox("�ϴ�����ʧ�ܣ��Ƿ�ֹͣ�����ϴ���ֱ�ӽ��н��ʣ�", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                            intMsg = 1
                            Exit Do
                        Else
                            intMsg = -1
                        End If
                    End If
                Else
                    '�����ϴ��ɹ������ϱ��
                    gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & rs��ϸ("ID") & "," & Val(arrOutput(2)) * rs��ϸ("����") & ",'" & arrOutput(1) & "')"
                    '�������ط����ϴ���ͬ��û�в�����һ�����Ӵ�ִ�С���Ϊ�������������õ���һ��ع���
                    cn�ϴ�.Execute gstrSQL, , adCmdStoredProc
                End If
            Else
                '�����ϴ��ɹ������ϱ��
                gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & rs��ϸ("ID") & ")"
                '�������ط����ϴ���ͬ��û�в�����һ�����Ӵ�ִ�С���Ϊ�������������õ���һ��ع���
                cn�ϴ�.Execute gstrSQL, , adCmdStoredProc
            End If
            
            rs��ϸ.MoveNext
        Loop
    End If
    
    '����Ԥ����
    Call DebugTool("׼������Ԥ����")
    strInput = "06|" & GetIdentify(lng����ID, g��������.��ҳID)
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    cur�����ʻ� = Val(arrOutput(2))
    curͳ��֧�� = Val(arrOutput(1))
    cur���ͳ�� = Val(arrOutput(5))
    cur����Ա���� = Val(arrOutput(3))
    
    '���没�˸����ʻ����
    mstrҽ���� = strҽ����
    mdbl��� = cur�����ʻ�
    
    '������ʱ���ݣ�Ϊ���������׼��
    With g��������
        .�������ý�� = cur��������
    End With
    
    str�ܽ��ҽԺ = Format(cur��������, "#####0.00")
    str�ܽ��ҽ�� = Format(curͳ��֧�� + cur�����ʻ� + cur����Ա���� + cur���ͳ�� + Val(arrOutput(4)), "#####0.00")
    If str�ܽ��ҽԺ <> str�ܽ��ҽ�� Then
        If MsgBox("ҽԺ�ķ����ܽ��(" & str�ܽ��ҽԺ & ")��ҽ�����ĵķ����ܶ�(" & str�ܽ��ҽ�� & ")���ȣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    סԺ�������_���� = "ҽ������;" & curͳ��֧�� & ";0"
    If cur�����ʻ� <> 0 Then
        סԺ�������_���� = סԺ�������_���� & "|�����ʻ�;" & cur�����ʻ� & ";0" '�������޸ĸ����ʻ�
    End If
    If cur���ͳ�� <> 0 Then
        סԺ�������_���� = סԺ�������_���� & "|���ͳ��;" & cur���ͳ�� & ";0"
    End If
    If cur����Ա���� <> 0 Then
        סԺ�������_���� = סԺ�������_���� & "|����Ա����;" & cur����Ա���� & ";0"
    End If
    If Val(arrOutput(6)) > 0 Then
        סԺ�������_���� = סԺ�������_���� & "|����Ա����;" & Val(arrOutput(6)) & ";0"
    End If
    
    '����Ԥ������ڽ���ʱ�ٱȽ�һ�Σ�������ֲ��
    With g��������
        .ͳ�ﱨ����� = curͳ��֧��       '1
        .�����ʻ�֧�� = cur�����ʻ�       '2
        .�ۼƽ���ͳ�� = cur����Ա����     '3
        .ȫ�Էѽ�� = Val(arrOutput(4))   '4
        .����ͳ���� = cur���ͳ��       '5
        .�ۼ�ͳ�ﱨ�� = Val(arrOutput(6)) '6
    End With
    
    mlng����id = lng����ID  '��ʾ�ò����Ѿ��������������
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long, ByVal lng����ID As Long) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
'      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    
    On Error GoTo errHandle
    Dim rsTemp As New ADODB.Recordset, strInput As String, arrOutput  As Variant
    Dim str����Ա As String, lng�����־ As Long
    Dim curͳ��֧�� As Double, cur�����ʻ� As Double
    Dim cur���ͳ�� As Double, cur����Ա���� As Double
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, datCurr As Date, strNO As String
    Dim strFormat As String
    Dim str���� As String, str����֢ As String, str������Ϣ As String, str��ע As String
    
    If mlng����id <> lng����ID Then
        MsgBox "�ò���û�����ҽ����Ԥ������������ܽ��н��㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHandle
    Call DebugTool("����סԺ����")
    '��ȡ�α����˵Ĳ���֢��������Ϣ
    gstrSQL = "Select ����֤�� As ���ֱ���,��������,����֢,������Ϣ" & _
        " From �����ʻ� " & _
        " Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ�α����˵Ĳ���֢��������Ϣ")
    If Not rsTemp.EOF Then
        str���� = Nvl(rsTemp!���ֱ���)
        If str���� <> "" Then str���� = "[" & str���� & "]"
        str���� = str���� & Nvl(rsTemp!��������)
        str����֢ = Nvl(rsTemp!����֢)
        str������Ϣ = Nvl(rsTemp!������Ϣ)
    End If
    str��ע = str���� & "||" & str����֢ & "||" & str������Ϣ
    
    '����Ԥ����
    Call DebugTool("׼�����ý���")
    strInput = "06|" & GetIdentify(lng����ID, g��������.��ҳID)
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    '�����ؽ�����Ԥ�������һ�αȽ�
    strFormat = "#####0.00;-#####0.00;0;"
    With g��������
        If CDbl(Format(.ͳ�ﱨ�����, strFormat)) <> CDbl(Format(arrOutput(1), strFormat)) Or _
           CDbl(Format(.�����ʻ�֧��, strFormat)) <> CDbl(Format(arrOutput(2), strFormat)) Or _
           CDbl(Format(.�ۼƽ���ͳ��, strFormat)) <> CDbl(Format(arrOutput(3), strFormat)) Or _
           CDbl(Format(.ȫ�Էѽ��, strFormat)) <> CDbl(Format(arrOutput(4), strFormat)) Or _
           CDbl(Format(.����ͳ����, strFormat)) <> CDbl(Format(arrOutput(5), strFormat)) Or _
           CDbl(Format(.�ۼ�ͳ�ﱨ��, strFormat)) <> CDbl(Format(arrOutput(6), strFormat)) Then
            
           If MsgBox("����������Ԥ����Ľ����һ�£����������ڲ��������µķ����ϴ�������������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End With
    
    '������ʻ�֧�����
    gstrSQL = "Select Nvl(��Ԥ��,0) as ��� From ����Ԥ����¼ Where ���㷽ʽ='�����ʻ�' And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "סԺ����")
    If Not rsTemp.EOF Then cur�����ʻ� = rsTemp("���")
    
    '���ý���
    With g��������
        If .��;���� = 1 Then
'            If MsgBox("�ò����Ƿ����ת��ͥ�������㣿", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
'                lng�����־ = 20 '��Ժת��ͥ����
'            Else
                lng�����־ = 10 '��;����
'            End If
        Else
            lng�����־ = 0 '��������
        End If
        
        strInput = "05|" & GetIdentify(lng����ID, .��ҳID) & "|" & lng�����־ & "|" & g��������.סԺ���� & "|" & UserInfo.���� & "|0" '�ø����ʻ����֧��
    End With
    
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    '��д�����
    Call DebugTool("��д�����¼")
    datCurr = zlDatabase.Currentdate
    curͳ��֧�� = Val(arrOutput(2))
    cur����Ա���� = Val(arrOutput(4))
    cur���ͳ�� = Val(arrOutput(6))
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
            
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & _
        cur����ͳ���ۼ� + curͳ��֧�� & "," & _
        curͳ�ﱨ���ۼ� + curͳ��֧�� & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure("����ҽ��")
    
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,NULL,0," & g��������.�������ý�� & ",0,0," & _
        curͳ��֧�� & "," & curͳ��֧�� & ",0,0,0,'" & arrOutput(1) & "'," & g��������.��ҳID & "," & g��������.��;���� & ",'" & str��ע & "')"
    Call ExecuteProcedure("����ҽ��")
    
    '���ս������
    gstrSQL = "zl_���ս������_insert(" & lng����ID & ",0," & curͳ��֧�� & "," & curͳ��֧�� & ",NULL)"
    Call ExecuteProcedure("����ҽ��")
    
    '��Ժ����ʱ����סԺ����ŵ�������Ϊ��
    If g��������.��;���� = 0 Then
        Dim strסԺ����� As String
        If GetMode(lng����ID, g��������.��ҳID, strסԺ�����) = False Then
            '��סԺ�������Ϊ�գ�Ϊ�´�סԺ��׼��
            gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'סԺ�����','''" & "" & "''')"
            Call ExecuteProcedure("��סԺ�������Ϊ��")
        End If
    End If
    
    סԺ����_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����(lng����ID As Long) As Boolean
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
    Dim intסԺ�����ۼ� As Integer
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
    If CanסԺ�������(rsTemp("����ID"), rsTemp("��ҳID")) = False Then Exit Function
    
    str��ˮ�� = rsTemp("֧��˳���")
    
    strInput = "99|" & str��ˮ�� & "|" & ToVarchar(UserInfo.����, 20)
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(rsTemp("����ID"), Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & rsTemp("����ID") & "," & gintInsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - rsTemp("�����ʻ�֧��") & "," & cur����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure("����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & rsTemp("����ID") & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - rsTemp("�����ʻ�֧��") & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & rsTemp("�������ý��") * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0,0," & _
        rsTemp("�����ʻ�֧��") * -1 & ",'" & str��ˮ�� & "'," & rsTemp("��ҳID") & "," & rsTemp("��;����") & ")"
    Call ExecuteProcedure("����ҽ��")

    סԺ�������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ������Ϣ_����(ByVal lngErrCode As Long) As String
'���ܣ����ݴ���ŷ��ش�����Ϣ

End Function

Public Function ҽԺ����_����() As String
'���ܣ��õ�ҽԺ��ҽ������
    Dim strInput As String, arrOutput As Variant
    
    On Error GoTo errHandle
    
    strInput = "11"
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    ҽԺ����_���� = arrOutput(1)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function HandleBusiness(ByVal strInput As String, varOut As Variant) As Boolean
'���ܣ�����ҽ������������ҵ����
    Dim strReturn As String '����ǰ�÷������ķ���ֵ
    Dim lngReturn As Long
    Dim varArray As Variant, lngCount As Long
    
    On Error Resume Next
    varOut = ""
    Screen.MousePointer = vbHourglass
    strReturn = Space(1024)
    lngReturn = dy_Business_Handle(strInput, strReturn)
    If Err <> 0 Or lngReturn = -1 Then
        varArray = Split(strReturn, "|")
        
        If UBound(varArray) > 0 Then
            strReturn = "ҽ���ӿڵ���ʧ�ܡ�" & vbCrLf & varArray(1)
        Else
            strReturn = "ҽ���ӿڵ���ʧ�ܡ�" & vbCrLf & strReturn
        End If
        Screen.MousePointer = vbDefault
        MsgBox strReturn, vbExclamation, gstrSysName
        Exit Function
    End If
    strReturn = TruncZero(strReturn)
    
    varArray = Split(strReturn, "|")
    If varArray(0) = "-1" Then
        'ҵ�����ʧ��
        If UBound(varArray) > 0 Then
            strReturn = "ҽ���ӿڳ��־��档" & vbCrLf & varArray(1)
        Else
            strReturn = "ҽ��ҵ����ʧ�ܡ�"
        End If
        
        Screen.MousePointer = vbDefault
        MsgBox strReturn, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '���׳ɹ�
    varOut = Split(strReturn, "|")
    
    HandleBusiness = True
    Screen.MousePointer = vbDefault
End Function

Private Function Get���ղ���_����(ByVal str������ As String) As String
'���ܣ���ñ��ղ���
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.������,A.����ֵ from ���ղ��� A " & _
              " where A.������='" & str������ & "' and A.����=" & TYPE_������ & " and A.���� is null "
    Call OpenRecordset(rsTemp, "����ҽ��")
    
    If rsTemp.EOF = False Then
        Get���ղ���_���� = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
    End If
End Function

Public Function �۸��ж�_����(ByVal dblҽԺ As Double, ByVal dblҽ�� As Double, ByVal str�޼۷�ʽ As String, _
                              ByVal bln�ؼ� As Boolean, ByVal dbl�ؼ� As Double) As Boolean
'���ܣ��ж�ҽԺ�ļ۸��Ƿ񳬹�ҽ���涨�ĵ���
    Dim strҽԺ��� As String
    
    On Error GoTo errHandle
    
    If InStr(str�޼۷�ʽ, "����") > 0 Then
        strҽԺ��� = Get���ղ���_����("ҽԺ�ȼ�")
        '�����ı�׼�۸�Ϊ����ҽԺ������޼ۣ�����ҽԺ������޼��ڴ˻����Ͽ����ϸ�10%��һ��ҽԺ������޼��ڴ˻������µ�5%
        
        Select Case strҽԺ���
            Case "����"
                dblҽ�� = dblҽ�� * 1.1
            Case "һ��"
                dblҽ�� = dblҽ�� * 0.95
        End Select
    End If
    
    If bln�ؼ� = True And dbl�ؼ� > dblҽ�� Then
        '����ʹ���ؼ�
        dblҽ�� = dbl�ؼ�
    End If
    
    If dblҽԺ > dblҽ�� Then
        If MsgBox("ҽԺ����" & Format(dblҽԺ, "0.000") & " ����ҽ�����ĺ�׼�ļ۸�" & Format(dblҽ��, "0.000") & "���Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    �۸��ж�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ���ʴ���_����(ByVal str���ݺ� As String, ByVal int���� As Integer, str��Ϣ As String, Optional ByVal lng����ID As Long = 0) As Boolean
'����:�ϴ��²����ļ�����ϸ��ҽ������
'����:  str���ݺ�   NO
'       int����     ��¼����
'       str��Ϣ    �����������������ѣ�����ǰ̨������ɣ����ⳤʱ���������
'       lng����ID  Ĭ��Ϊ0����ʾ�������ŵ��ݣ�����Ϊ������ָ�����˵ġ�����Ҫ����Ϊҽ���ڱ�����ʵ�ʱ���Ƿֲ������ύ���ݶ�����һ���ύ��
'����:
    Dim rsTemp As New ADODB.Recordset
    Dim rsTest As New ADODB.Recordset
    Dim cn�ϴ� As New ADODB.Connection
    Dim strInput As String, arrOutput, arrTemp  As Variant, curͳ���� As Currency
    Dim strҽ�� As String, str������ As String
    Dim col���� As New Collection, lngPre����ID As Long, var���� As Variant, bln�ɹ� As Boolean
    Dim str������ϸ��ˮ�� As String, str������Ϣ As String
    '��ע�⣺����ҽ�����ڼ��ʵ�������ٵ��ô�����̵ġ�
    
    On Error GoTo errHandle
    
    cn�ϴ�.ConnectionString = gcnOracle.ConnectionString
    cn�ϴ�.Open
    
    '�������ŵ��ݵķ�����ϸ
    gstrSQL = "Select A.ID,A.NO,A.����ID,A.��ҳID,A.����ʱ�� as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս�� " & _
              "         ,A.�շ�ϸĿID,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
              "         ,C.��Ŀ����,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,A.����Ա����,B.���㵥λ,E.���,G.���� ���� " & _
              "  From ���˷��ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C,������ҳ D,ҩƷĿ¼ E ,ҩƷ��Ϣ F,ҩƷ���� G " & _
              "  where A.NO='" & str���ݺ� & "' and A.��¼����=" & int���� & " and A.��¼״̬=1 And Nvl(A.�Ƿ��ϴ�,0)=0 " & _
              "        and A.����ID=D.����ID and A.��ҳID=D.��ҳID And D.����=" & gintInsure & IIf(lng����ID = 0, "", " and A.����ID=" & lng����ID) & _
              "        and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����=D.���� " & _
              "        AND B.ID=E.ҩƷID(+) AND E.ҩ��ID=F.ҩ��ID(+) AND F.����=G.����(+) " & _
              "  Order by A.����ID,A.����ʱ��"
    Call OpenRecordset(rsTemp, "���ʴ���")
    
    Do While Not rsTemp.EOF
        If Val(rsTemp!����) < 0 Or Val(rsTemp!�۸�) < 0 Then
            '����ȡһ��������¼����ˮ�ţ���Ϊ������ˮ��
            str������ϸ��ˮ�� = GetSequence(rsTemp!����ID, rsTemp!��ҳID, rsTemp!�շ�ϸĿID)
            If Trim(str������ϸ��ˮ��) = "" Then
                MsgBox "û���ҵ����Գ����ļ�¼��[" & rsTemp!���� & "]" & rsTemp!����, vbInformation, gstrSysName
                Exit Function
            End If
        Else
            str������ϸ��ˮ�� = ""
        End If
        rsTemp.MoveNext
    Loop
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    
    '����ҽ����Ŀ����Ϣ������
    Do While Not rsTemp.EOF
        Call TrackRecordInsure(rsTemp!ID, rsTemp!�շ�ϸĿID)
        rsTemp.MoveNext
    Loop
    
    '���з�����ϸ�Ĵ���
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do Until rsTemp.EOF
        If Val(rsTemp!����) < 0 Or Val(rsTemp!�۸�) < 0 Then
            '����ȡһ��������¼����ˮ�ţ���Ϊ������ˮ��
            str������ϸ��ˮ�� = GetSequence(rsTemp!����ID, rsTemp!��ҳID, rsTemp!�շ�ϸĿID)
        Else
            str������ϸ��ˮ�� = ""
        End If
        
        strҽ�� = ToVarchar(IIf(IsNull(rsTemp("ҽ��")), UserInfo.����, rsTemp("ҽ��")), 20)
        str������ = ToVarchar(IIf(IsNull(rsTemp("����Ա����")), UserInfo.����, rsTemp("����Ա����")), 20)
        
        strInput = "04|" & GetIdentify(rsTemp("����ID"), rsTemp("��ҳID"))
        strInput = strInput & "|" & rsTemp("NO") & "_" & int����
        strInput = strInput & "|" & Format(rsTemp("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")
        strInput = strInput & "|" & ToVarchar(rsTemp("��Ŀ����"), 10)     '���ı���
        strInput = strInput & "|" & ToVarchar(rsTemp("����"), 20)         'ҽԺ����
        strInput = strInput & "|" & ToVarchar(rsTemp("����"), 50)         '��Ŀ����
        strInput = strInput & "|" & Format(rsTemp("�۸�"), "0.0000")      '����
        strInput = strInput & "|" & Format(rsTemp("����"), "0.00")        '����
        strInput = strInput & "|" & IIf(rsTemp("�Ƿ���") = 1, 1, 0)     '�����־
        strInput = strInput & "|" & strҽ��                               'ҽ��
        strInput = strInput & "|" & str������                             '������
        strInput = strInput & "|" & ToVarchar(rsTemp("���㵥λ"), 20)     '��λ
        strInput = strInput & "|" & ToVarchar(rsTemp("���"), 14)         '���
        strInput = strInput & "|" & ToVarchar(rsTemp("����"), 20)         '����
        strInput = strInput & "|" & str������ϸ��ˮ��                     '������ϸ��ˮ��
        strInput = strInput & "|" & Format(rsTemp("ʵ�ս��"), "#####0.0000")         '���
        
        If HandleBusiness(strInput, arrOutput) = False Then
            If bln�ɹ� = True Then
                MsgBox "�����ϴ���;�������󣬲����Ѿ������Ѿ��ϴ�������Ԥ���㴦���ʣ�����ݵ��ϴ���", vbInformation, gstrSysName
            Else
                MsgBox "�����ϴ���������û�гɹ��ϴ��ļ�¼������Ԥ���㴦���ʣ�����ݵ��ϴ���", vbInformation, gstrSysName
            End If
            ���ʴ���_���� = True
            Exit Function
        End If
        Call AddMessage(str��Ϣ, arrOutput, rsTemp("����"), rsTemp("�۸�")) '���Բ�����������Ϣ
        
        If lngPre����ID <> rsTemp("����ID") Then '�ж�ʱû�п�����ҳID������Ϊͬһ���˲�����ͬʱ������סԺ����ϸ
            'Modified by ZYB 2004-05-10
            '��ȡ���˵Ļ�����Ϣ��������ڷ���ԭ������ʾ
            gstrSQL = "Select ҽ���� From �����ʻ� Where ����=" & TYPE_������ & " And ����ID=" & rsTemp!����ID
            Call OpenRecordset(rsTest, "ȡ�ò��˵�ҽ����")
            
            strInput = "01|" & rsTest!ҽ����
            If HandleBusiness(strInput, arrTemp) Then
                str������Ϣ = ""
                If Val(arrTemp(11)) <> 0 Then
                    str������Ϣ = arrTemp(12)
                    MsgBox str������Ϣ, vbInformation, gstrSysName
                End If
                '���·�����Ϣ
                gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'������Ϣ','''" & str������Ϣ & "''')"
                Call ExecuteProcedure("���²���")
            End If
            
            '���Ѿ��ϴ��Ĳ�����Ϣ��¼��������Ϊ���ʱ��Ƕಡ�˵ģ�
            col����.Add rsTemp("����ID") & "_" & rsTemp("��ҳID")
            lngPre����ID = rsTemp("����ID")
        End If
        
        '�ڷ��ü�¼�ϴ��ϱ�ǣ�˵���Ѿ��ϴ��������淵�صĽ��
        If arrOutput(3) = 2 Then
            'δͨ������
            curͳ���� = 0
        Else
            '��׼���� * ����
            curͳ���� = Val(arrOutput(2)) * rsTemp("����")
        End If
        gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & rsTemp("ID") & "," & curͳ���� & ",'" & arrOutput(1) & "')"
        cn�ϴ�.Execute gstrSQL, , adCmdStoredProc
        bln�ɹ� = True
        
        rsTemp.MoveNext
    Loop
    
    If str��Ϣ <> "" Then
        str��Ϣ = "���˷�����ϸ��������еõ�ҽ���������·�����Ϣ����Ŀǰ�����Ѿ����档" & vbCrLf & "����кβ��ף������ѡ�����ϸõ��ݡ�" & vbCrLf & vbCrLf & str��Ϣ
    End If
        
    ���ʴ���_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��������_����(ByVal str���ݺ� As String, ByVal int���� As Integer, str��Ϣ As String) As Boolean
'����:�����Ѿ��ϴ���ҽ�����ĵļ�����ϸ
'����:  str���ݺ�   NO
'       int����     ��¼����
'       str��Ϣ    �����������������ѣ�����ǰ̨������ɣ����ⳤʱ���������
'����:
    
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, arrOutput As Variant
    Dim lngPre����ID As Long
    Dim strҽ�� As String, str������ As String, str������ϸ��ˮ�� As String
    Dim bln�ɹ� As Boolean
    Dim cn�ϴ� As New ADODB.Connection
    
    On Error GoTo errHandle
    
    cn�ϴ�.ConnectionString = gcnOracle.ConnectionString
    cn�ϴ�.Open
    
    '�������ŵ��ݵķ�����ϸ����δ�ϴ��ļ�¼��ȡԭʼ���ݣ�
    gstrSQL = "Select nvl(count(A.ID),0) as ����,nvl(sum(A.�Ƿ��ϴ�),0) �ϴ��� " & _
              "  From ���˷��ü�¼ A,������ҳ B,����֧����Ŀ C" & _
              "  where A.NO='" & str���ݺ� & "' and A.��¼����=" & int���� & " and A.��¼״̬<>2 And Nvl(A.��¼״̬,0)<>0 And Nvl(A.�Ƿ��ϴ�,0)=0 and nvl(A.ʵ�ս��,0)<>0  " & _
              "        and A.����ID=B.����ID and A.��ҳID=B.��ҳID And B.����=" & gintInsure & " and A.�շ�ϸĿID=C.�շ�ϸĿID and B.����=C.����"
    Call OpenRecordset(rsTemp, "��������")
    
    If rsTemp.EOF = True Then
        MsgBox "�õ�����û�п��ϴ������Ϸ�����ϸ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If rsTemp("�ϴ���") = 0 Then
        '��ϸ������û���ϴ�������Ҳ�Ͳ���Ҫ��������
        ��������_���� = True
        Exit Function
    End If
    
    If rsTemp("�ϴ���") < rsTemp("����") Then
        MsgBox "�õ����ﻹ��δ�ϴ��ķ�����ϸ���������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�����õ����ڲ����������Ϊ���ʱ��Ƕಡ�˵ģ�
    gstrSQL = " Select A.ID,A.�շ�ϸĿID,A.NO,A.��¼����,A.��¼״̬,A.���,A.����ID,A.��ҳID,A.����ʱ�� as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս��" & _
              "         ,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
              "         ,C.��Ŀ����,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,A.����Ա����,B.���㵥λ,E.���,G.���� ���� " & _
              "  From ���˷��ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C,������ҳ D,ҩƷĿ¼ E ,ҩƷ��Ϣ F,ҩƷ���� G " & _
              "  where A.NO='" & str���ݺ� & "' and A.��¼����=" & int���� & " and A.��¼״̬=2 and nvl(A.�Ƿ��ϴ�,0)=0" & _
              "        and A.����ID=D.����ID and A.��ҳID=D.��ҳID And D.����=" & gintInsure & _
              "        and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����=D.���� " & _
              "        AND B.ID=E.ҩƷID(+) AND E.ҩ��ID=F.ҩ��ID(+) AND F.����=G.����(+) " & _
              "  Order by A.����ʱ��,A.��¼����,Decode(A.��¼״̬,2,2,1)"
    Call OpenRecordset(rsTemp, "��������")
    
    '�Ƚ�ҽ����Ŀ��Ϣ��������������
    Do While Not rsTemp.EOF
        Call TrackRecordInsure(rsTemp!ID, rsTemp!�շ�ϸĿID)
        rsTemp.MoveNext
    Loop
    
    '���з�����ϸ�Ĵ���
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do Until rsTemp.EOF
        '���ʱ�ͼ��ʵ��������ֳ�������Ҫһ��һ�ʳ���
        str������ϸ��ˮ�� = GetDetailSequence(rsTemp!NO, rsTemp!���, rsTemp!��¼����, rsTemp!��¼״̬)
        str������ = ToVarchar(IIf(IsNull(rsTemp("����Ա����")), UserInfo.����, rsTemp("����Ա����")), 20)
        strInput = "99|" & str������ϸ��ˮ�� & "|" & str������
        If HandleBusiness(strInput, arrOutput) = False Then
            '�����ϴ�ʧ��
            If bln�ɹ� = True Then
                MsgBox "�����ϴ���;�������󣬲����Ѿ������Ѿ��ϴ�������Ԥ���㴦���ʣ�����ݵ��ϴ���", vbInformation, gstrSysName
            Else
                MsgBox "�����ϴ���������û�гɹ��ϴ��ļ�¼������Ԥ���㴦���ʣ�����ݵ��ϴ���", vbInformation, gstrSysName
            End If
            ��������_���� = True
            Exit Function
        Else
            '�ڲ��������Ϸ��ü�¼�ϴ��ϱ�ǣ�˵���Ѿ��ϴ�
            gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & rsTemp("ID") & ")"
            '������һ�����Ӵ�ִ�У��ѳɹ��ϴ��Ĵ����ϴ���־
            cn�ϴ�.Execute gstrSQL, , adCmdStoredProc
        End If
        
        rsTemp.MoveNext
        bln�ɹ� = True
    Loop
    
    ��������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub AddMessage(strMessage As String, arrOutput As Variant, ByVal str��Ŀ As String, ByVal dbl���� As Currency)
'���ܣ��ڲ��˷�����ϸ��������п��ܲ���һЩ��Ҫ���Ѳ�����Ա����Ϣ
    Dim strTemp As String
    
    If dbl���� > Val(arrOutput(2)) And Val(arrOutput(2)) > 0 Then
        strTemp = "��    " & str��Ŀ & "��ҽԺ�۸� " & Format(dbl����, "0.0000") & " �������ķ��ؼ۸� " & Format(Val(arrOutput(2)), "0.0000") & vbCrLf
    End If
    If arrOutput(3) = 2 Then
        strTemp = "��    " & str��Ŀ & "��Ҫ��������û��������¼��ֻ����Ϊ�Է���Ŀ" & vbCrLf
    End If
    
    If InStr(strMessage, strTemp) = 0 Then
        strMessage = strMessage & strTemp
    End If
    
End Sub

Private Function GetDetailSequence(ByVal strNO As String, ByVal int��� As Integer, _
        ByVal int���� As Integer, ByVal int״̬ As Integer) As String
    Dim rsTemp As New ADODB.Recordset
    '����NO����¼���ʡ���¼״̬�����ȡԭʼ��¼����ˮ�ţ����ڳ�����
    GetDetailSequence = ""
    If int״̬ <> 2 Then Exit Function
    
    gstrSQL = " Select ժҪ From ���˷��ü�¼" & _
              " Where NO='" & strNO & "' And ���=" & int��� & _
              " And ��¼����=" & int���� & " And ��¼״̬=3"
    Call OpenRecordset(rsTemp, "ȡԭʼ������ϸ����ˮ��")
    If Not rsTemp.EOF Then
        GetDetailSequence = Nvl(rsTemp!ժҪ)
    Else
        Call DebugTool("δ�ҵ�ԭʼ������ϸ[NO:" & strNO & "|���:" & int��� & "|��¼����:" & int���� & "|��¼״̬:" & int״̬)
    End If
End Function

Private Function GetSequence(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng�շ�ϸĿID As Long) As String
    '���ȡһ��������¼����ˮ�ţ����ڸ������ʣ�
    Dim rsTemp As New ADODB.Recordset
    GetSequence = ""
    
    gstrSQL = " Select ժҪ From ���˷��ü�¼" & _
              " Where �շ�ϸĿID=" & lng�շ�ϸĿID & " And ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID & _
              " And ��¼״̬=1 And Nvl(�Ƿ��ϴ�,0)=1 And Nvl(ʵ�ս��,0)>0 And Rownum<2"
    Call OpenRecordset(rsTemp, "ȡ��ˮ��")
    If Not rsTemp.EOF Then
        GetSequence = rsTemp!ժҪ
    Else
        Call DebugTool("δ�ҵ�ԭʼ������ϸ[����ID:" & lng����ID & "|��ҳID:" & lng��ҳID & "|�շ�ϸĿID:" & lng�շ�ϸĿID)
    End If
End Function

Private Function TrackRecordInsure(ByVal lng����ID As Long, ByVal lng�շ�ϸĿID As Long) As Boolean
    Dim str��ˮ�� As String, str�������� As String, str�շ���� As String, str��ע As String
    Dim dbl��׼���� As Double, dbl�Ը����� As Double
    Dim rsTemp As New ADODB.Recordset
    '��¼ҽ����Ŀ��ʱ�Ļ�����Ϣ��ҽ����Ŀ���룬�������ͣ���׼����,�Ը�������
    Call DebugTool("����TrackRecordInsure")
    gstrSQL = "Select A.���,B.��Ŀ���� " & _
        " From �շ�ϸĿ A,����֧����Ŀ B" & _
        " Where A.ID=B.�շ�ϸĿID And B.����=" & gintInsure & " And A.ID=" & lng�շ�ϸĿID
    Call OpenRecordset(rsTemp, "��ȡ����Ŀ���������")
    If rsTemp.RecordCount = 0 Then Exit Function
    str��ˮ�� = Nvl(rsTemp!��Ŀ����)
    str�շ���� = rsTemp!���
    Call DebugTool("��ǰҽ����Ŀ���룺" & str��ˮ��)
    If str��ˮ�� = "" Then Exit Function
    
    '������
    If mcnYB.State = 0 Then
        If Not OpenDatabase Then Exit Function
    End If
    
    If InStr(1, "5,6,7", str�շ����) <> 0 Then
        'ҩƷ
        gstrSQL = "select YPLSH  ҽ������,YPBM ҩƷ����,TYM ͨ������,SPM ��Ʒ��,SPMZJM ��Ʒ������,YCMC ҩ������,decode(FYDJ,1,'����',2,'����','�Է�') ���õȼ� " & _
                  "      ,PFJ ������,BZDJ ��׼����,ZFBL �Ը�����,JX ����,BZSL ��װ����,BZDW ��װ��λ,HL ����,HLDW ������λ,RL ����,RLDW ������λ " & _
                  "      ,DECODE(CFYBZ,1,'��') ����ҩ��־,decode(GMP,1,'��') GMP��־,decode(YPXJFS,1,'�޼�') �޼�,TQFYDJ ��Ⱥ��Ŀ�ȼ�,TQZFBL ��Ⱥ�Ը�����,TQBZDJ ��Ⱥ��׼���� " & _
                  "  FROM YPML WHERE YPLSH='" & str��ˮ�� & "'"
    Else
        '����
        gstrSQL = "Select XMLSH ҽ������,XMBM ���Ʊ���,XMMC ��Ŀ����,ZJM ����,decode(FYDJ,1,'����',2,'����','�Է�') ���õȼ�,DW ��λ " & _
                 "       ,TPJ ������,BZJ ��׼����,ZZBL ��ְ�Ը�����,TXBL �����Ը�����,decode(XJFS,1,'ͳһ�޼�',2,'��ҽԺ�ȼ�����',3,'������ҽԺ��׼��������') �޼� " & _
                 "       ,TQFYDJ ��Ⱥ��Ŀ�ȼ�,TQZFBL ��Ⱥ�Ը�����,TQBZDJ ��Ⱥ��׼����,decode(TPXMBZ,1,'��') ������Ŀ��־,BZ ��ע " & _
                 "   FROM ZLXM WHERE XMLSH='" & str��ˮ�� & "'"
    End If
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, mcnYB
        If .RecordCount = 0 Then
            Call DebugTool("δ�ҵ���ҽ����Ŀ")
            Exit Function
        End If
    End With
    
    str�������� = Nvl(rsTemp!���õȼ�)
    dbl��׼���� = Nvl(rsTemp!��׼����, 0)
    If InStr(1, "5,6,7", str�շ����) <> 0 Then
        dbl�Ը����� = Nvl(rsTemp!�Ը�����, 0)
        str��ע = "||||" & Nvl(rsTemp!��Ⱥ�Ը�����, 0)
    Else
        dbl�Ը����� = Nvl(rsTemp!��ְ�Ը�����, 0)
        str��ע = Nvl(rsTemp!������, 0) & "||" & Nvl(rsTemp!�����Ը�����, 0) & "||" & Nvl(rsTemp!��Ⱥ�Ը�����, 0)
    End If
    
    '�����¼���������жϣ�������ڼ�¼��������£�������룩
    '����ID,ҽ����Ŀ����,��������,��׼����,�Ը�����,��ע
    gstrSQL = "zl_ҽ����Ŀ��Ϣ_INSERT(" & lng����ID & ",'" & str��ˮ�� & "','" & str�������� & "'," & _
        dbl��׼���� & "," & dbl�Ը����� & ",'" & str��ע & "')"
    Call DebugTool("����ҽ����Ŀ��Ϣ��" & gstrSQL)
    Call ExecuteProcedure("����ҽ����Ŀ��Ϣ��¼")
    TrackRecordInsure = True
End Function

Private Function OpenDatabase() As Boolean
    Dim strServer As String, strUser As String, strPass As String, strTemp As String
    Dim rsTemp As New ADODB.Recordset
    '���ȶ���������������
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & gintInsure
    Call OpenRecordset(rsTemp, "��ȡ���ղ���")
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "ҽ��������"
                strServer = strTemp
            Case "ҽ���û���"
                strUser = strTemp
            Case "ҽ���û�����"
                strPass = strTemp
        End Select
        rsTemp.MoveNext
    Loop
    If OraDataOpen(mcnYB, strServer, strUser, strPass) = False Then
        Exit Function
    End If
    OpenDatabase = True
End Function

Public Function �ҺŽ���_����(lng����ID As Long) As Boolean
    Dim intTotal As Integer, intStart As Integer
    Dim str���㷽ʽ As String, arr���㷽ʽ
    Dim cur�����ʻ� As Currency, curҽ������ As Currency, cur����Ա���� As Currency, cur���ͳ�� As Currency, cur����Ա���� As Currency
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    gstrSQL = "Select ����ID,�շ�ϸĿID,����*NVL(����,1) AS ����,��׼���� As ����,ʵ�ս��," & IIf(g��������.�����Ը���� = 14, "1", "0") & " As �Ƿ���" & _
        " From ���˷��ü�¼ " & _
        " Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Call OpenRecordset(rsTemp, "����ҽ��")
    
    If Not �����������_����(rsTemp, str���㷽ʽ) Then Exit Function
    If Not �������_����(lng����ID, 0, "") Then Exit Function
    
    '�ֽ���ֽ��㷽ʽ
    arr���㷽ʽ = Split(str���㷽ʽ, "|")
    intTotal = UBound(arr���㷽ʽ)
    For intStart = 0 To intTotal
        Select Case Split(arr���㷽ʽ(intStart), ";")(0)
        Case "�����ʻ�"
            cur�����ʻ� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        Case "ҽ������"
            curҽ������ = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        Case "����Ա����"
            cur����Ա���� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        Case "���ͳ��"
            cur���ͳ�� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        Case "����Ա����"
            cur����Ա���� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        End Select
    Next
    
   '��Ҫ����������
    str���㷽ʽ = ""
    If cur�����ʻ� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & cur�����ʻ�
    If curҽ������ <> 0 Then str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & curҽ������
    If cur���ͳ�� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||���ͳ��|" & cur���ͳ��
    If cur����Ա���� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & cur����Ա����
    If cur����Ա���� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & cur����Ա����
    If str���㷽ʽ <> "" Then
        str���㷽ʽ = Mid(str���㷽ʽ, 3)
    Else
        str���㷽ʽ = "�����ʻ�|0"
    End If
    gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',0)"
    Call ExecuteProcedure("����Ԥ����¼")
    
    �ҺŽ���_���� = True
    Call frm������Ϣ.ShowME(lng����ID)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetIdentify(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
    Dim strסԺ����� As String
    'MODIFIED BY ZYB 20040626 '��Ժ�޸ģ�����������ϵͳ�ǰ�����ID�ϴ��ģ���ˣ���ˮ�Ÿ�Ϊ��
    '��������ʻ��д���סԺ����ţ���ֵ��Ϊ�գ���������Ϊ���˱�ʶ�ϴ�����������ģʽ�ϴ�
    If GetMode(lng����ID, lng��ҳID, strסԺ�����) Then
        GetIdentify = lng����ID & "_" & lng��ҳID
    Else
        GetIdentify = strסԺ�����
    End If
End Function

Private Function GetMode(ByVal lng����ID As Long, ByVal lng��ҳID As Long, strסԺ����� As String) As Boolean
    'MODIFIED BY ZYB 20040626 '��Ժ�޸ģ�����������ϵͳ�ǰ�����ID�ϴ��ģ���ˣ���ˮ�Ÿ�Ϊ��
    '��������ʻ��д���סԺ����ţ���ֵ��Ϊ�գ���������Ϊ���˱�ʶ�ϴ�����������ģʽ�ϴ�
    Dim blnģʽ As Boolean              'Ϊ�棬������ID_��ҳID���أ���������ز���ID
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    gstrSQL = " Select סԺ����� From �����ʻ� " & _
              " Where ����=" & TYPE_������ & " ANd ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "ȡסԺ�����")
    If Err <> 0 Then
        '˵�������ڸ��ֶ�
        blnģʽ = True
    Else
        blnģʽ = (Nvl(rsTemp!סԺ�����) = "")
        If Not blnģʽ Then strסԺ����� = Nvl(rsTemp!סԺ�����)
    End If
    GetMode = blnģʽ
End Function
