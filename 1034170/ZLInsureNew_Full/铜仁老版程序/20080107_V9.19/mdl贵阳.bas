Attribute VB_Name = "mdl����"
Option Explicit

Public mdomInput As MSXML2.DOMDocument
Public mdomOutput As MSXML2.DOMDocument

Private mstr���� As String
Private mstr���� As String

Private mstrҽ���� As String
Private mdbl��� As Double

Private mlng����ID As Long

Public Function ҽ����ʼ��_����() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    
    On Error Resume Next
    
    Set mdomInput = New MSXML2.DOMDocument
    If Err <> 0 Then
        MsgBox "���ܴ���XML����������ע��msxml3.dll������", vbInformation, gstrSysName
    Else
        ҽ����ʼ��_���� = True
    End If
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim str������� As String, str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, cur�ʻ���� As Currency
    Dim str���� As String, str�Ա� As String, str���֤���� As String, lng���� As Long
    Dim str�������� As String, str��Ա��� As String, str��λ���� As String, str��λ���� As String
    Dim strIdentify As String, str���� As String, lng����ID As Long
    Dim rsTemp As New ADODB.Recordset, rs���� As ADODB.Recordset
    
    '��ʼ��һЩ�������ڳ�����;�˳�ʱֵȴ�Ѿ�����
    mstr���� = "": mstr���� = ""
    If frmIdentify����.GetIdentify(TYPE_������, str����, strҽ����, str�����ı��, str����, True, True) = False Then
        Exit Function
    End If
    '��ԭ����
    str������� = Split(str����, "^")(1)
    str���� = Split(str����, "^")(0)
    
    If bytType = id����ȷ�� Then
        '�÷���ֵ��ʱû�����ã�ֻҪ��Ϊ�վͱ�ʾ�ɹ���
        ��ݱ�ʶ_���� = str���� & ";" & strҽ���� & ";" & str����
        Exit Function
    End If
    
    'ȡ�÷���ֵ
    str���� = GetElemnetValue("PERSONNAME")
    str�Ա� = GetElemnetValue("SEX")
    str�Ա� = Switch(str�Ա� = "1", "��", str�Ա� = "2", "Ů", str�Ա� = "9", "����", True, str�Ա�)
    str���֤���� = GetElemnetValue("PID")
    
    str�������� = AddDate(GetElemnetValue("BIRTHDAY"))
    If IsDate(str��������) = True Then
        lng���� = DateDiff("yyyy", CDate(str��������), zlDatabase.Currentdate)
    Else
        str�������� = ""
    End If
    
    str��Ա��� = GetElemnetValue("PERSONTYPE")
    str��Ա��� = Switch(str��Ա��� = "11", "��ְ", str��Ա��� = "21", "����" _
                      , str��Ա��� = "32", "ʡ������", str��Ա��� = "34", "��������", True, "����")
    str��λ���� = ToVarchar(GetElemnetValue("DEPTCODE"), 12)
    str��λ���� = ToVarchar(GetElemnetValue("DEPTNAME"), 36) '�ֶγ��ȱ���50�������ڻ�Ҫ������뼰����
    cur�ʻ���� = Val(GetElemnetValue("ACCTBALANCE"))
    
    
    '����;ҽ����;����;����;�Ա�;��������;���֤;������λ
    'ҽ���ŵ�һλΪ������
    strIdentify = str���� & ";" & strҽ���� & ";" & str���� & ";" & str���� & ";" & str�Ա� & ";" & str�������� & ";" & str���֤���� & ";" & str��λ���� & "(" & str��λ���� & ")"
    strIdentify = Replace(strIdentify, " ", "")
    
    '��������
    'Modified By ���� 2003-12-03 ������ ԭ����Ժʱȡ������ѡ�񣬸�Ϊ���������ʱ�����û�в��֣�����ѡ��
    If bytType = id�����շ� And Get���ղ���_����("֧����������") = "1" Then
        gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                " From ���ղ��� A where A.����=" & gintInsure
        
        Set rs���� = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ������")
        If Not rs���� Is Nothing Then
            lng����ID = rs����("ID")
        End If
    End If
    
    str���� = ";"                                       '8.���Ĵ���
    str���� = str���� & ";" & str�����ı��             '9.˳���  ����ҽ�����ڱ���ҽ�������ı��루���⽨��ҽ�����ģ�
    str���� = str���� & ";" & str��Ա���               '10��Ա���
    str���� = str���� & ";" & cur�ʻ����               '11�ʻ����
    str���� = str���� & ";0"                            '12��ǰ״̬
    str���� = str���� & ";" & IIf(lng����ID <> 0, lng����ID, "")   '13����ID
    str���� = str���� & ";" & IIf(str��Ա��� = "��ְ", 1, 2)      '14��ְ(1,2)
    str���� = str���� & ";"                             '15����֤��
    str���� = str���� & ";" & lng����                   '16�����
    str���� = str���� & ";"                             '17�Ҷȼ�
    str���� = str���� & ";" & cur�ʻ����               '18�ʻ������ۼ�
    str���� = str���� & ";0"                            '19�ʻ�֧���ۼ�
    str���� = str���� & ";"                             '20����ͳ���ۼ�
    str���� = str���� & ";"                             '21ͳ�ﱨ���ۼ�
    str���� = str���� & ";"                             '22סԺ�����ۼ�
    str���� = str���� & ";"                             '23�������� (1����������)
    
    lng����ID = BuildPatiInfo(bytType, strIdentify & str����, lng����ID)
    '���ظ�ʽ:�м���벡��ID
    If lng����ID <> 0 Then
        ��ݱ�ʶ_���� = strIdentify & ";" & lng����ID & str����
        
        mstr���� = str����
        mstr���� = str����
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��������_������(ByVal str�ſ����� As String, ByVal str���� As String, ByVal str������ As String) As Boolean
    If InitXML = False Then Exit Function
    
    Call InsertChild(mdomInput.documentElement, "CARDDATA", str�ſ�����)            ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str����)                ' ����
    Call InsertChild(mdomInput.documentElement, "NEWPASSWORD", str������)           ' ����
    
    '���ýӿ�
    If CommServer("MODIFYCARD") = False Then Exit Function
    ��������_������ = True
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
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, str��Ա��� As String
    Dim dbl�����ʻ� As Double
    Dim lng����ID As Long, str�������� As String, datCurr As Date
    Dim rsTemp As New ADODB.Recordset
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    On Error GoTo errHandle
    
    If rs��ϸ.RecordCount = 0 Then
        str���㷽ʽ = "�����ʻ�;0;0"
        �����������_���� = True
        Exit Function
    End If
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ("����ID")
    
    '�жϸò����Ƿ�����������
    gstrSQL = "select A.��Ա���,B.���� from �����ʻ� A,���ղ��� B where A.����ID=" & lng����ID & " and A.����=" & gintInsure & "  and A.����ID=B.ID(+)"
    Call OpenRecordset(rsTemp, "����Ԥ��")
    If rsTemp.EOF = False Then
        str�������� = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        str��Ա��� = Nvl(rsTemp("��Ա���"), "")
        'ת����Ա���
        str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21" _
                      , str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", True, "11")
    End If
    datCurr = zlDatabase.Currentdate
    
    If Get��֤_����(str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDID", str����)           ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", strҽ����)     ' ���˱���
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str�����ı��) ' �����ı���
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str����)         ' ����
    Call InsertChild(mdomInput.documentElement, "PERSONTYPE", str��Ա���)         ' ����
    If str�������� <> "" Then '��������
        '����8λ����
        str�������� = String(8 - Len(str��������), "0") & str��������
        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str��������)         '���ֲ�����
        Call InsertChild(mdomInput.documentElement, "STARTDATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) '������ʼ����ʱ��
    End If
    Call InsertChild(mdomInput.documentElement, "ISCAL", 0)         ' �Ƿ����
    Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", "0")     ' �˻�֧����
    Call InsertChild(mdomInput.documentElement, "INVOICENO", " ") ' ��Ʊ��
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) ' ��������
    Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' ������ϸ
    
    Do Until rs��ϸ.EOF
        gstrSQL = "SELECT C.ҩƷID,C.���,E.���� AS ����  FROM ҩƷĿ¼ C,ҩƷ��Ϣ D,ҩƷ���� E WHERE C.ҩƷID=" & rs��ϸ("�շ�ϸĿID") & " AND C.ҩ��ID=D.ҩ��ID AND D.����=E.����"
        gstrSQL = "select A.���,A.����,B.��Ŀ����,nvl(A.���,F.���) AS ���,F.����,A.���㵥λ from �շ�ϸĿ A,����֧����Ŀ B,(" & gstrSQL & _
                ") F where A.ID=" & rs��ϸ("�շ�ϸĿID") & " and A.ID=B.�շ�ϸĿID  AND A.Id=F.ҩƷID(+) and B.����=" & gintInsure
        Call OpenRecordset(rsTemp, "����Ԥ��")
        If rsTemp.EOF = True Then
            MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
            Exit Function
        End If

        Set nodRow = InsertChild(nodRowset, "ROW", "")
        Call nodRow.setAttribute("ITEMCODE", ToVarchar(rsTemp("��Ŀ����"), 12))
        Call nodRow.setAttribute("ITEMNAME", ToVarchar(rsTemp("����"), 72))
        Call nodRow.setAttribute("SUBJECT", Subject(rsTemp("���")))
        Call nodRow.setAttribute("SPECIFICATION", ToVarchar(rsTemp("���"), 40))
        Call nodRow.setAttribute("AGENTTYPE", ToVarchar(rsTemp("����"), 20))
        Call nodRow.setAttribute("UNIT", ToVarchar(rsTemp("���㵥λ"), 20))
        Call nodRow.setAttribute("PRICE", Format(rs��ϸ("����"), "0.0000"))
        Call nodRow.setAttribute("QUANTITY", Format(rs��ϸ("����"), "0.00"))
        Call nodRow.setAttribute("FROMOFFICE", ToVarchar(UserInfo.����, 56)) '��������
        Call nodRow.setAttribute("FROMDOCT", Format(UserInfo.����, 20))      '����ҽ��
        Call nodRow.setAttribute("TOOFFICE", ToVarchar(UserInfo.����, 56))  '�ܵ�����
        Call nodRow.setAttribute("TODOCT", Format(UserInfo.����, 20))       '�ܵ�ҽ��
        Call nodRow.setAttribute("DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss"))        '��������
        Call nodRow.setAttribute("NOTE", ToVarchar(rs��ϸ("ժҪ"), 512))        '��ע
        
        rs��ϸ.MoveNext
    Loop
    
    '���ýӿ�
    If CommServer(IIf(str�������� <> "", "CALSPECCLIN", "CALCLIN")) = False Then Exit Function
    
    '������Ա��������ͨ�������������ͳһ��ALLOWFUND֧����
    '����ҽ����Ա����������FUND1PAY��FUND2PAY֧������ͨ�����ɸ����ʻ�֧��
    If str��Ա��� = "32" Or str��Ա��� = "34" Then
        str���㷽ʽ = "ҽ������;" & Val(GetElemnetValue("ALLOWFUND")) & ";0"
    Else
        str���㷽ʽ = "�����ʻ�;" & Val(GetElemnetValue("ACCTPAY")) & ";1"  '�����޸ĸ����ʻ�
        If str�������� <> "" Then
            str���㷽ʽ = str���㷽ʽ & "|ҽ������;" & Val(GetElemnetValue("FUND1PAY")) & ";0" & _
                         "|��ͳ��;" & Val(GetElemnetValue("FUND2PAY")) & ";0"
        End If
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
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, str��Ա��� As String
'    Dim str����Re As String, strҽ����Re As String, str�����ı��Re As String, str����Re As String
    Dim strҽ�� As String, str���� As String, cur�������� As Double, datCurr As Date
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim lng����ID  As Long, str��������   As String, lng��Ŀ�� As Long
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    On Error GoTo errHandle
    lng��Ŀ�� = Val(Get���ղ���_����("���������Ŀ��"))

    gstrSQL = "SELECT Nvl(��������,Nvl(�۸񸸺�,���)) AS ����� FROM ���˷��ü�¼  " & _
             " WHERE ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0" & _
             " GROUP BY Nvl(��������,Nvl(�۸񸸺�,���))"
    Call OpenRecordset(rs��ϸ, "����ҽ��")
    If rs��ϸ.RecordCount > lng��Ŀ�� Then
        MsgBox "�����շѵ���Ŀ�����ܳ���" & lng��Ŀ�� & "��", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "Select A.ID,A.���,A.����ID,A.NO,A.�Ǽ�ʱ��,A.������ as ҽ��,A.�Ǽ�ʱ��," & _
            "   A.����*A.���� as ����,A.��׼���� as ʵ�ʼ۸�,A.���ʽ��," & _
            "   A.�շ����,D.��Ŀ����,B.���� as ��Ŀ����,C.���� as ��������,nvl(B.���,F.���) AS ���,F.����,B.���㵥λ,A.ժҪ " & _
            " From (Select * From ���˷��ü�¼ Where ����ID=" & lng����ID & ") A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D " & _
            "     ,(SELECT C.ҩƷID,C.���,E.���� AS ����  FROM ���˷��ü�¼ A,ҩƷĿ¼ C,ҩƷ��Ϣ D,ҩƷ���� E WHERE A.����ID=" & lng����ID & " AND A.�շ�ϸĿID=C.ҩƷID AND C.ҩ��ID=D.ҩ��ID AND D.����=E.����) F " & _
            " Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID And A.�շ�ϸĿID=D.�շ�ϸĿID  AND B.ID=F.ҩƷID(+) And D.����=" & gintInsure & " And Nvl(A.���ӱ�־,0)<>9 And Nvl(A.��¼״̬,0)<>0" & _
            " Order by A.ID"
    Call OpenRecordset(rs��ϸ, "����ҽ��")
    If rs��ϸ.EOF = True Then
        MsgBox "û����д�շѼ�¼", vbExclamation, gstrSysName
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")
    strҽ�� = ToVarchar(IIf(IsNull(rs��ϸ("ҽ��")), UserInfo.����, rs��ϸ("ҽ��")), 20)
    str���� = ToVarchar(IIf(IsNull(rs��ϸ("��������")), UserInfo.����, rs��ϸ("��������")), 56)
    datCurr = zlDatabase.Currentdate
    
    
    'һ��������ϸ����
    
    '�жϸò����Ƿ�����������
    gstrSQL = "select A.��Ա���,B.���� from �����ʻ� A,���ղ��� B where A.����ID=" & lng����ID & " and A.����=" & gintInsure & "  and A.����ID=B.ID(+)"
    Call OpenRecordset(rsTemp, "����Ԥ��")
    If rsTemp.EOF = False Then
        str�������� = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        str��Ա��� = Nvl(rsTemp("��Ա���"), "")
        'ת����Ա���
        str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21" _
                      , str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", True, "11")
    End If
    
    If Get��֤_����(str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
    '�����շ�ʱ��Ҫ��ˢһ�ο�
'    If frmIdentify����.GetIdentify(TYPE_������, str����Re, strҽ����Re, str�����ı��Re, str����Re, False) = False Then
'        Exit Function
'    Else
'        If str���� <> str����Re Or strҽ���� <> strҽ����Re Then
'            MsgBox "��ʹ�õ�ǰ���˵Ŀ���ˢһ�Ρ�", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
        
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDID", str����)           ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", strҽ����)     ' ���˱���
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str�����ı��) ' �����ı���
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str����)         ' ����
    Call InsertChild(mdomInput.documentElement, "PERSONTYPE", str��Ա���)         ' ����
    If str�������� <> "" Then '��������
        '����8λ����
        str�������� = String(8 - Len(str��������), "0") & str��������
        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str��������)         '���ֲ�����
        Call InsertChild(mdomInput.documentElement, "STARTDATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) '������ʼ����ʱ��
    End If
    Call InsertChild(mdomInput.documentElement, "ISCAL", 1)         ' �Ƿ����
    Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", Format(cur�����ʻ�, "0.00")) ' �˻�֧����
    Call InsertChild(mdomInput.documentElement, "INVOICENO", rs��ϸ("NO")) ' ��Ʊ��
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(rs��ϸ("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")) ' ��������
    Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' ������ϸ
    
    Do Until rs��ϸ.EOF
        cur�������� = cur�������� + rs��ϸ("���ʽ��")
        
        Set nodRow = InsertChild(nodRowset, "ROW", "")
        Call nodRow.setAttribute("ITEMCODE", ToVarchar(rs��ϸ("��Ŀ����"), 12))
        Call nodRow.setAttribute("ITEMNAME", ToVarchar(rs��ϸ("��Ŀ����"), 72))
        Call nodRow.setAttribute("SUBJECT", Subject(rs��ϸ("�շ����")))
        Call nodRow.setAttribute("SPECIFICATION", ToVarchar(rs��ϸ("���"), 40))
        Call nodRow.setAttribute("AGENTTYPE", ToVarchar(rs��ϸ("����"), 20))
        Call nodRow.setAttribute("UNIT", ToVarchar(rs��ϸ("���㵥λ"), 20))
        Call nodRow.setAttribute("PRICE", Format(rs��ϸ("ʵ�ʼ۸�"), "0.0000"))
        Call nodRow.setAttribute("QUANTITY", Format(rs��ϸ("����"), "0.00"))
        Call nodRow.setAttribute("FROMOFFICE", str����)    '��������
        Call nodRow.setAttribute("FROMDOCT", strҽ��)      '����ҽ��
        Call nodRow.setAttribute("TOOFFICE", str����)     '�ܵ�����
        Call nodRow.setAttribute("TODOCT", strҽ��)       '�ܵ�ҽ��
        
        '����ʱ��ʱ��Ϊ�˱�֤ͬһ������Ŀ�ĵ��շ�ʱ�䲻ͬ������ڵǼ�ʱ���ϰ���ż�������
        Call nodRow.setAttribute("DODATE", Format(rs��ϸ("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss"))    '��������
        Call nodRow.setAttribute("NOTE", ToVarchar(rs��ϸ("ժҪ"), 512))         '��ע
        
        rs��ϸ.MoveNext
    Loop
    
    '���ýӿ�
    If CommServer(IIf(str�������� <> "", "CALSPECCLIN", "CALCLIN")) = False Then Exit Function
    
    
    '��������¼
    '---------------------------------------------------------------------------------------------
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
            
    Dim curȫ�Ը� As Double, cur�ҹ��Ը� As Double, curͳ��֧�� As Double
    Dim curͳ���Ը� As Double, cur�����Ը� As Double, cur�����Ը� As Double
    Dim cur��ͳ�� As Double, cur���Ը� As Double, cur���� As Double
    Dim str����˳��� As String, str������ As String
            
    curȫ�Ը� = Val(GetElemnetValue("FEEOUT"))
    cur�ҹ��Ը� = Val(GetElemnetValue("FEESELF"))
    cur���� = Val(GetElemnetValue("STARTFEE"))
    cur�����Ը� = Val(GetElemnetValue("ENTERSTARTFEE"))
    If str��Ա��� = "32" Or str��Ա��� = "34" Then
        curͳ��֧�� = Val(GetElemnetValue("ALLOWFUND"))
        cur��ͳ�� = 0
    Else
        curͳ��֧�� = Val(GetElemnetValue("FUND1PAY"))
        cur��ͳ�� = Val(GetElemnetValue("FUND2PAY"))
    End If
    curͳ���Ը� = Val(GetElemnetValue("FUND1SELF"))
    cur���Ը� = Val(GetElemnetValue("FUND2SELF"))
    cur�����Ը� = Val(GetElemnetValue("FEEOVER"))
    
    str������ = GetElemnetValue("BALANCEID")
    str����˳��� = GetElemnetValue("BILLNO")
    If str�������� <> "" Then
        str����˳��� = "����" & str�������� & str����˳��� '�Ѽ������������˳�������һ��
    Else
        str����˳��� = "��ͨ" & str����˳���         '��ʾ��ͨ����
    End If
    
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
                
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� + curͳ��֧�� + curͳ���Ը� + cur�����Ը� + cur�����Ը� + cur��ͳ�� + cur���Ը� & "," & _
        curͳ�ﱨ���ۼ� + curͳ��֧�� + cur��ͳ�� & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure("����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & ",0," & cur�����Ը� & "," & cur�������� & "," & _
        curȫ�Ը� & "," & cur�ҹ��Ը� & "," & _
        curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & "," & cur���Ը� & "," & cur�����Ը� & "," & cur�����ʻ� & ",'" & str������ & "',null,null,'" & str����˳��� & "')"
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
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    Dim str������ As String, str����˳��� As String, curDate As Date, rsTemp As New ADODB.Recordset
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency, intסԺ�����ۼ� As Integer
    Dim lng����ID As Long
    
    On Error GoTo errHandle
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "�����˷�")
    lng����ID = rsTemp("����ID")
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=" & gintInsure & " and ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, "�����˷�")
    If rsTemp.EOF = True Then
        MsgBox "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    lng����ID = rsTemp!����ID
    cur�����ʻ� = Nvl(rsTemp!�����ʻ�֧��, 0)
    str������ = Nvl(rsTemp("֧��˳���"), "")
    str����˳��� = Nvl(rsTemp("��ע"), "")
    If str����˳��� = "" Then
        MsgBox "�õ���û�б������˳��ţ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If Left(str����˳���, 2) = "����" Then
        MsgBox "Ŀǰ��֧��������������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    str����˳��� = Mid(str����˳���, 3)
    curDate = zlDatabase.Currentdate
    
    If Get��֤_����(str����, strҽ����, str�����ı��, str����, lng����ID, True) = False Then Exit Function
        
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDID", str����)           ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", strҽ����)     ' ���˱���
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str�����ı��) ' �����ı���
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str����)         ' ����
    Call InsertChild(mdomInput.documentElement, "BILLNO", str����˳���)     ' ����˳���
    Call InsertChild(mdomInput.documentElement, "BALANCEID", str������)    ' ������
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)    ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(curDate, "yyyy-MM-dd HH:mm:ss"))  ' ��������
    
    '���ýӿ�
    If CommServer("RETCLIN") = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure("����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & rsTemp("�������ý��") * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0," & rsTemp("�����Ը����") & "," & _
        cur�����ʻ� * -1 & ",'" & str������ & "',null,null,'" & Nvl(rsTemp("��ע"), "") & "')"
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
    Dim str���� As String, str�����ı�� As String, str���� As String, str��Ա��� As String
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str�������� As String
    Dim strtemp As String, str��ʾ As String, str��� As String, lng�α�ǰ��Ժ As Long
    
    On Error GoTo errHandle
    
    If Get��֤_����(str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
    
    '�жϸò����Ƿ�α�ǰ��Ժ
    lng�α�ǰ��Ժ = 0
    If Get���ղ���_����("��Ժʱѡ��α�ǰ��Ժ") = "1" Then
        If MsgBox("�ò��˲α�ǰ�Ƿ��Ѿ���Ժ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            lng�α�ǰ��Ժ = 1
        End If
    End If
    
    '�жϸò����Ƿ������ⲡ
    gstrSQL = "select A.��Ա���,B.���� from �����ʻ� A,���ղ��� B where A.����ID=" & lng����ID & " and A.����=" & gintInsure & "  and A.����ID=B.ID(+)"
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    If rsTemp.EOF = False Then
        str�������� = Nvl(rsTemp("����"), "")
        str��Ա��� = Nvl(rsTemp("��Ա���"), "")
        'ת����Ա���
        str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21" _
                      , str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", True, "11")
    End If
    
    '��ò��˳�Ժ���
    gstrSQL = "select A.������Ϣ from ������ A where A.����ID=" & lng����ID & " and A.��ҳID=" & lng��ҳID & _
              " and A.�������=1 and A.��ϴ���=1"
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    If rsTemp.EOF = False Then
        str��� = ToVarchar(IIf(IsNull(rsTemp("������Ϣ")), "����", rsTemp("������Ϣ")), 128)
    Else
        str��� = "����"   '��ϲ�����β���Ϊ��
    End If
    
    '���������Ժ��Ϣ
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select A.��Ժ��ʽ,nvl(A.����Ժת��,0) as ����Ժת��,A.����ҽʦ,A.��Ժ����,A.��Ժ����,B.���� as ��Ժ����,C.סԺ�� from ������ҳ A,���ű� B,������Ϣ C " & _
              " Where A.����ID=C.����ID and A.��Ժ����ID = B.ID And A.����ID =" & lng����ID & " And A.��ҳID = " & lng��ҳID
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDID", str����)           ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", strҽ����)     ' ���˱���
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str�����ı��) ' �����ı���
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str����)         ' ����
    Call InsertChild(mdomInput.documentElement, "PERSONTYPE", str��Ա���)   ' ��Ա���
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", IIf(rsTemp("��Ժ��ʽ") = "ת��", "37", "31"))     ' ֧����� 31��סԺ��37��תԺ
    
    If str�������� <> "" Then
        '����8λ����
        str�������� = String(8 - Len(str��������), "0") & str��������
        Call InsertChild(mdomInput.documentElement, "FUND2FLAG", "1")                 ' ת����־
        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str��������)   ' ���ֲ�����
    Else
        'û�����ⲡ
        Call InsertChild(mdomInput.documentElement, "FUND2FLAG", "")            ' ת����־
        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", "")      ' ���ֲ�����
    End If
    
    Call InsertChild(mdomInput.documentElement, "HOSPNO", ToVarchar(rsTemp("סԺ��"), 20))     ' סԺ��
    Call InsertChild(mdomInput.documentElement, "ISINHOSP", lng�α�ǰ��Ժ)     ' �α�ǰ����Ժ 1���ǣ�0����
    Call InsertChild(mdomInput.documentElement, "DIAGNOSES", str���) ' ���
    Call InsertChild(mdomInput.documentElement, "DOCTOR", ToVarchar(rsTemp("����ҽʦ"), 20)) ' ���ҽ��
    Call InsertChild(mdomInput.documentElement, "OFFICE", ToVarchar(rsTemp("��Ժ����"), 20)) ' ����
    Call InsertChild(mdomInput.documentElement, "POSITION", ToVarchar(rsTemp("��Ժ����"), 10)) ' ��λ
    Call InsertChild(mdomInput.documentElement, "REGDATE", Format(rsTemp("��Ժ����"), "yyyy-MM-dd HH:mm:ss")) ' ��Ժʱ��
    Call InsertChild(mdomInput.documentElement, "OPERATOR", UserInfo.����) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss"))  ' ��������
    
    '���ýӿ�
    If CommServer("HOSPREG") = False Then Exit Function
    
    Dim intסԺ�����ۼ� As Integer
    Dim cur�������� As Currency
    Dim cur�����ۼ� As Currency
    Dim cur����ͳ���޶� As Currency
    Dim curͳ�ﱨ���ۼ� As Currency
    Dim cur���ͳ���޶� As Currency
    Dim cur���ͳ���ۼ� As Currency
    
    Dim str������Ϣ As String
    
    intסԺ�����ۼ� = Val(GetElemnetValue("HOSPTIMES"))
    
    cur�������� = Val(GetElemnetValue("STARTFEE"))
    cur�����ۼ� = Val(GetElemnetValue("STARTFEEPAID"))
    cur����ͳ���޶� = Val(GetElemnetValue("FUND1LMT"))
    curͳ�ﱨ���ۼ� = Val(GetElemnetValue("FUND1PAID"))
    cur���ͳ���޶� = Val(GetElemnetValue("FUND2LMT"))
    cur���ͳ���ۼ� = Val(GetElemnetValue("FUND2PAID"))
    
    str������Ϣ = GetElemnetValue("LOCKINFO")
    Do Until str������Ϣ = ""
        strtemp = Left(str������Ϣ, 2)
        str������Ϣ = Mid(str������Ϣ, 41)
        
        str��ʾ = str��ʾ & Switch(strtemp = "11", "�����������", strtemp = "21", "��������", strtemp = "31", "������ͳ��Ƿ��", _
                                   strtemp = "32", "�����ͳ��δ�ɷ�", strtemp = "41", "��ͣ��", strtemp = "51", "���˱�")
        
    Loop
    If str��ʾ <> "" Then
        MsgBox "��ע���ҽ�����������" & Mid(str��ʾ, 2) & "��", vbInformation, gstrSysName
    End If
    
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        "0,0,0," & curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�������� & "," & cur�����ۼ� & _
         "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & "," & cur���ͳ���ۼ� & ",'" & ToVarchar(str��ʾ, 100) & "')"
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
    Dim strҽ���� As String, str�����ı�� As String
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str��� As String, str������� As String
    Dim str������ As String, str��Ժת�� As String, lngPos As Long
    
    On Error GoTo errHandle
    
    '�����ݿ��ж����Ѵ洢��ֵ
    gstrSQL = "select ����,ҽ����,˳��� from �����ʻ� where ����ID=" & lng����ID & " and ����=" & gintInsure
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    
    strҽ���� = IIf(IsNull(rsTemp("ҽ����")), "", rsTemp("ҽ����"))
    str�����ı�� = IIf(IsNull(rsTemp("˳���")), "", rsTemp("˳���"))
    
    '��ò��˳�Ժ��Ϣ
    gstrSQL = "SELECT A.��Ժ��ʽ,nvl(C.������,B.סԺ��) AS ������  " & _
             " FROM ������ҳ A,������Ϣ B,סԺ������¼ C " & _
             " WHERE A.����ID=" & lng����ID & " AND A.��ҳid=" & lng��ҳID & " AND A.����id=B.����id AND A.����id=C.����id(+)"
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    str������ = rsTemp("������")
    Select Case rsTemp("��Ժ��ʽ")
        Case "����", "����"
            str��Ժת�� = "1"
        Case "��ת"
            str��Ժת�� = "2"
        Case "����"
            str��Ժת�� = "3"
        Case Else
            str��Ժת�� = "9"
    End Select
    
    '��ò��˳�Ժ���
    gstrSQL = "select A.������Ϣ from ������ A where A.����ID=" & lng����ID & " and A.��ҳID=" & lng��ҳID & _
              " and A.�������=3 and A.��ϴ���=1"
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    If rsTemp.EOF = False Then
        str��� = Nvl(rsTemp("������Ϣ"), "����")
        '����ͬ��ʽ�ķָ���ͳһ
        str��� = Replace(str���, "��", ",")
        str��� = Replace(str���, "��", ",")
        str��� = Replace(str���, "��", ",")
        str��� = Replace(str���, ";", ",")
        lngPos = InStr(str���, ",")
        If lngPos > 0 Then
            str������� = Mid(str���, lngPos + 1)
            str��� = Mid(str���, 1, lngPos - 1)
        End If
    Else
        str��� = "����"   '��ϲ�����β���Ϊ��
    End If
        
    '���������Ժ��Ϣ
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select A.סԺҽʦ,A.��Ժ����,A.��Ժ����,B.���� as ��Ժ���� from ������ҳ A,���ű� B " & _
             " Where A.��Ժ����ID = B.ID And A.����ID =" & lng����ID & " And A.��ҳID = " & lng��ҳID
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", strҽ����)     ' ���˱���
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str�����ı��) ' �����ı���
    Call InsertChild(mdomInput.documentElement, "DOCNO", str������)          ' ������
    Call InsertChild(mdomInput.documentElement, "DIAGNOSES", ToVarchar(str���, 128))          ' ���
    Call InsertChild(mdomInput.documentElement, "OTHERDIAGNOSES", ToVarchar(str�������, 128)) ' �������
    Call InsertChild(mdomInput.documentElement, "OUTTYPE", str��Ժת��)                        ' ת�����
    Call InsertChild(mdomInput.documentElement, "DOCTOR", ToVarchar(rsTemp("סԺҽʦ"), 20))   ' ���ҽ��
    Call InsertChild(mdomInput.documentElement, "OFFICE", ToVarchar(rsTemp("��Ժ����"), 20))   ' ����
    'Call InsertChild(mdomInput.documentElement, "POSITION", ToVarchar(rsTemp("��Ժ����"), 10)) ' ��λ
    Call InsertChild(mdomInput.documentElement, "REGDATE", Format(rsTemp("��Ժ����"), "yyyy-MM-dd HH:mm:ss")) ' ��Ժ����
    Call InsertChild(mdomInput.documentElement, "OPERATOR", UserInfo.����) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss"))  ' ��������
    
    '���ýӿ�
    If CommServer("HOSPOUT") = False Then Exit Function
    
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


Public Function סԺ�������_����(rsExse As Recordset, ByVal lng����ID As Long) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim cn�ϴ� As New ADODB.Connection, rsTemp As New ADODB.Recordset, rs���� As ADODB.Recordset
    Dim lng����ID As Long, str�������� As String
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, str��Ա��� As String
    Dim cur�����ʻ� As Double, curͳ��֧�� As Double, cur��ͳ�� As Double, cur�������� As Double
    Dim strҽ�� As String, str���� As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    On Error GoTo errHandle
    mlng����ID = 0         '��ʼ����ֻҪһѡ���ˣ��ͻ���ñ����̣�Ҳ�ͻ����0
    
    If rsExse.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If
    
    rsExse.MoveFirst
    '������һ�����Ӵ����Դﵽ���ܵ�ǰ��������Ŀ���
    cn�ϴ�.ConnectionString = gcnOracle.ConnectionString
    cn�ϴ�.Open
    
    '�˴�����ȷ���ǵò����ģ�����Ҫǿ��ˢ��
    Screen.MousePointer = vbDefault
    
    'ȡ�ò��˵Ļ�����Ϣ
    gstrSQL = "select A.��Ա���,B.���� from �����ʻ� A,���ղ��� B where A.����ID=" & lng����ID & " and A.����=" & gintInsure & "  and A.����ID=B.ID(+)"
    Call OpenRecordset(rsTemp, "סԺԤ��")
    If rsTemp.EOF = False Then
        str��Ա��� = Nvl(rsTemp("��Ա���"), "")
        'ת����Ա���
        str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21" _
                      , str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", True, "11")
    End If
    
    mstr���� = ""
    mstr���� = ""
    If Get��֤_����(str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
    Screen.MousePointer = vbHourglass
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDID", str����)           ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", strҽ����)     ' ���˱���
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str�����ı��) ' �����ı���
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str����)         ' ����
    Call InsertChild(mdomInput.documentElement, "PERSONTYPE", str��Ա���)         ' ��Ա���
    Call InsertChild(mdomInput.documentElement, "ISCAL", 0)         ' �Ƿ����
    Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", "0")     ' �˻�֧����
    Call InsertChild(mdomInput.documentElement, "INVOICENO", " ") ' ��Ʊ��
'    'Modified By ���� 2003-12-03 ������ ԭ�������ϴ����ֱ���
'    '����8λ����
'    If str�������� <> "" Then
'        str�������� = String(8 - Len(str��������), "0") & str��������
'        Call InsertChild(mdomInput.documentElement, "FUND2FLAG", "1")                 ' ת����־
'        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str��������)   ' ���ֲ�����
'    End If
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) ' ��������
    Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' ������ϸ
    
    rsExse.Sort = "NO,���,�Ǽ�ʱ�� asc"
    Do Until rsExse.EOF
        If IIf(IsNull(rsExse("�Ƿ��ϴ�")), "0", rsExse("�Ƿ��ϴ�")) = "0" Then
            gstrSQL = "SELECT C.ҩƷID,C.���,E.���� AS ����  FROM ҩƷĿ¼ C,ҩƷ��Ϣ D,ҩƷ���� E WHERE C.ҩƷID=" & rsExse("�շ�ϸĿID") & " AND C.ҩ��ID=D.ҩ��ID AND D.����=E.����"
            gstrSQL = "select A.���,A.����,B.��Ŀ����,nvl(A.���,F.���) AS ���,F.����,A.���㵥λ from �շ�ϸĿ A,����֧����Ŀ B,(" & gstrSQL & _
                    ") F where A.ID=" & rsExse("�շ�ϸĿID") & " and A.ID=B.�շ�ϸĿID  AND A.Id=F.ҩƷID(+) and B.����=" & gintInsure
            Call OpenRecordset(rsTemp, "סԺԤ��")
            If rsTemp.EOF = True Then
                MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
                Exit Function
            End If
            
            'ֻ�ϴ�ֻ���ݹ�������
            strҽ�� = ToVarchar(IIf(IsNull(rsExse("ҽ��")), UserInfo.����, rsExse("ҽ��")), 20)
            str���� = ToVarchar(IIf(IsNull(rsExse("��������")), UserInfo.����, rsExse("��������")), 56)
            
            Set nodRow = InsertChild(nodRowset, "ROW", "")
            Call nodRow.setAttribute("ITEMSERIAL", ToVarchar(rsExse("NO") & "_" & rsExse("���") & "_" & rsExse("��¼����") & "_" & rsExse("��¼״̬"), 20)) '�������ţ�����Ψһ��������
            Call nodRow.setAttribute("ITEMCODE", ToVarchar(rsExse("ҽ����Ŀ����"), 12))
            Call nodRow.setAttribute("ITEMNAME", ToVarchar(rsExse("�շ�����"), 72))
            Call nodRow.setAttribute("SUBJECT", Subject(rsTemp("���")))
            Call nodRow.setAttribute("SPECIFICATION", ToVarchar(rsTemp("���"), 40))
            Call nodRow.setAttribute("AGENTTYPE", ToVarchar(rsTemp("����"), 20))
            Call nodRow.setAttribute("UNIT", ToVarchar(rsTemp("���㵥λ"), 20))
            Call nodRow.setAttribute("PRICE", Format(rsExse("�۸�"), "0.0000"))
            Call nodRow.setAttribute("QUANTITY", Format(rsExse("����"), "0.00"))
            Call nodRow.setAttribute("FROMOFFICE", str����)   '��������
            Call nodRow.setAttribute("FROMDOCT", strҽ��)     '����ҽ��
            Call nodRow.setAttribute("TOOFFICE", str����)    '�ܵ�����
            Call nodRow.setAttribute("TODOCT", strҽ��)      '�ܵ�ҽ��
            '����ʱ��ʱ��Ϊ�˱�֤ͬһ������Ŀ�ĵ��շ�ʱ�䲻ͬ������ڵǼ�ʱ���ϰ���ż�������
            Call nodRow.setAttribute("DODATE", Format(rsExse("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss"))      '��������
            Call nodRow.setAttribute("NOTE", ToVarchar(rsExse("ժҪ"), 512))     '��ע
        End If
        cur�������� = cur�������� + rsExse("���")
        rsExse.MoveNext
    Loop
    
    '���ýӿ�
    If CommServer("CALHOSP") = False Then Exit Function
    '����ǿ�������ٴ������Ե�ҽ����������ȷ���غ��ٴ��ϱ��
    If rsExse.RecordCount > 0 Then rsExse.MoveFirst
    Do Until rsExse.EOF
        If IIf(IsNull(rsExse("�Ƿ��ϴ�")), "0", rsExse("�Ƿ��ϴ�")) = "0" Then
            'Ϊ�������ü�¼�����ϴ���־���ϴ�һ������һ��
            gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsExse("NO") & "'," & rsExse("���") & "," & rsExse("��¼����") & "," & rsExse("��¼״̬") & ")"
            cn�ϴ�.Execute gstrSQL, , adCmdStoredProc
        End If
        rsExse.MoveNext
    Loop
    
    cur�����ʻ� = Val(GetElemnetValue("ACCTPAY"))
    If str��Ա��� = "32" Or str��Ա��� = "34" Then
        curͳ��֧�� = Val(GetElemnetValue("ALLOWFUND"))
    Else
        curͳ��֧�� = Val(GetElemnetValue("FUND1PAY"))
    End If
    cur��ͳ�� = Val(GetElemnetValue("FUND2PAY"))
    
    '���没�˸����ʻ����
    mstrҽ���� = strҽ����
    mdbl��� = cur�����ʻ�
    
    '������ʱ���ݣ�Ϊ���������׼��
    With g��������
        .�������ý�� = cur��������
    End With
    
    סԺ�������_���� = "ҽ������;" & curͳ��֧�� & ";0"
    If cur�����ʻ� <> 0 Then
        סԺ�������_���� = סԺ�������_���� & "|�����ʻ�;" & cur�����ʻ� & ";1" '�����޸ĸ����ʻ�
    End If
'    If cur��ͳ�� <> 0 Then
        '��������Ŀ���Ǳ���ǰ�˳����޸ĸý��㷽ʽ�Ľ��
        סԺ�������_���� = סԺ�������_���� & "|��ͳ��;" & cur��ͳ�� & ";0"
'    End If
    
    mlng����ID = lng����ID  '��ʾ�ò����Ѿ��������������
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
    Dim rsTemp As New ADODB.Recordset
    
    Dim curȫ�Ը� As Double, cur�ҹ��Ը� As Double, curͳ��֧�� As Double
    Dim curͳ���Ը� As Double, cur�����Ը� As Double, cur�����Ը� As Double
    Dim cur��ͳ�� As Double, cur���Ը� As Double, cur�����ʻ� As Double, cur���� As Currency
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, datCurr As Date, strNO As String
    Dim str����˳��� As String, str������ As String
    Dim lng����ID As Long
    Dim str�������� As String
    Dim rs���� As ADODB.Recordset
    
    If mlng����ID <> lng����ID Then
        MsgBox "�ò���û�����ҽ����Ԥ������������ܽ��н��㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    'Modified By ���� 2003-12-03 ������ ԭ����Ժ�Ǽ������֤��ȡ�����ֵ�ѡ�񣬸��ڽ���ʱ����ȷ������
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & gintInsure & " Order by A.����"
    
    Set rs���� = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ������")
    If Not rs���� Is Nothing Then
        lng����ID = rs����("ID")
        str�������� = rs����("����")
    Else
        lng����ID = 0
        str�������� = ""
    End If
    
    '���²�����Ϣ
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'����ID','" & lng����ID & "')"
    Call ExecuteProcedure("���²�����Ϣ")
    
    'Modified By ���� 2003-12-03 ������ ԭ�������ϴ����ֱ���
    '����8λ����
    If str�������� <> "" Then
        str�������� = String(8 - Len(str��������), "0") & str��������
        Call InsertChild(mdomInput.documentElement, "FUND2FLAG", "1")                 ' ת����־
        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str��������)   ' ���ֲ�����
    End If
    
    '������ʻ�֧�����
    gstrSQL = "Select Nvl(��Ԥ��,0) as ��� From ����Ԥ����¼ Where ���㷽ʽ='�����ʻ�' And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "סԺ����")
    If Not rsTemp.EOF Then cur�����ʻ� = rsTemp("���")
    '�󵥾ݺ�
    gstrSQL = "Select NO,�շ�ʱ�� From ���˽��ʼ�¼ Where ID=" & lng����ID
    Call OpenRecordset(rsTemp, "סԺ����")
    
    'XML�ĵ��Ѿ���ɳ�ʼ������ʱֻ��Ҫ���²���ֵ
    Call SetElemnetValue("ISCAL", "1")
    Call SetElemnetValue("ACCTWANTTOPAY", Format(cur�����ʻ�, "0.00"))
    Call SetElemnetValue("INVOICENO", rsTemp("NO"))
    Call SetElemnetValue("DODATE", Format(rsTemp("�շ�ʱ��"), "yyyy-MM-dd HH:mm:ss"))
    'Ԥ��ʱ�Ѿ����ݣ����ʲ���Ҫ�ٴ�����ϸ����
    Call SetElemnetValue("ROWSET", "")
    '���ýӿ�
    If CommServer("CALHOSP") = False Then Exit Function
    
    curȫ�Ը� = Val(GetElemnetValue("FEEOUT"))
    cur�ҹ��Ը� = Val(GetElemnetValue("FEESELF"))
    cur���� = Val(GetElemnetValue("STARTFEE"))
    cur�����Ը� = Val(GetElemnetValue("ENTERSTARTFEE"))
    curͳ��֧�� = Val(GetElemnetValue("FUND1PAY")) + Val(GetElemnetValue("ALLOWFUND"))
    curͳ���Ը� = Val(GetElemnetValue("FUND1SELF"))
    cur��ͳ�� = Val(GetElemnetValue("FUND2PAY"))
    cur���Ը� = Val(GetElemnetValue("FUND2SELF"))
    cur�����Ը� = Val(GetElemnetValue("FEEOVER"))
    
    str������ = GetElemnetValue("BALANCEID")
    str����˳��� = GetElemnetValue("BILLNO")
    
    '��д�����
    datCurr = zlDatabase.Currentdate
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
            
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� + curͳ��֧�� + curͳ���Ը� + cur�����Ը� + cur�����Ը� + cur��ͳ�� + cur���Ը� & "," & _
        curͳ�ﱨ���ۼ� + curͳ��֧�� + cur��ͳ�� & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure("����ҽ��")
    
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & ",NULL," & cur�����Ը� & "," & _
        g��������.�������ý�� & "," & curȫ�Ը� & "," & cur�ҹ��Ը� & "," & _
        curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & "," & cur���Ը� & "," & cur�����Ը� & "," & cur�����ʻ� & ",'" & str������ & "',null,null,'" & str����˳��� & "')"
    Call ExecuteProcedure("����ҽ��")
    
    '���ս������
    gstrSQL = "zl_���ս������_insert(" & lng����ID & ",0," & curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & ",NULL)"
    Call ExecuteProcedure("����ҽ��")
    
    '������㷽ʽ���ǰ����嵥����Ա�����������Ա������ʾ����ԱΪ�ò��˰����Ժ����
    gstrSQL = "Select ������,��Ա��� From �����ʻ� Where ����=" & TYPE_������ & " And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ���㷽ʽ")
    If Right(rsTemp!������, 1) <> 4 And Not (rsTemp!��Ա��� = "��������" Or rsTemp!��Ա��� = "ʡ������") Then
        MsgBox "��Ϊ�òα���Ա�����Ժ������", vbInformation, gstrSysName
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
    '      4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    Dim lng����ID As Long, lng����ID As Long
    Dim str�������� As String, str��ǰ���� As String
    Dim rsTemp  As New ADODB.Recordset, rsCheck As New ADODB.Recordset
    
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, str��Ա��� As String
    Dim str����˳��� As String, str������ As String
    Dim cur�����ʻ� As Currency
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim curDate As Date    '�˷�
    
    On Error GoTo ErrHand
    curDate = zlDatabase.Currentdate
    
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
    lng����ID = rsTemp!����ID
    str������ = IIf(IsNull(rsTemp!֧��˳���), "", rsTemp!֧��˳���)
    str����˳��� = IIf(IsNull(rsTemp!��ע), "", rsTemp!��ע)
    
    '�ж��Ƿ�Ϊ������Ա
    gstrSQL = "Select ��Ա��� From �����ʻ� Where ����ID=" & lng����ID & " And ����=" & gintInsure
    Call OpenRecordset(rsCheck, "�ж��Ƿ�Ϊ������Ա")
    If Not (rsCheck!��Ա��� = "ʡ������" Or rsCheck!��Ա��� = "��������") Then
        MsgBox "����ҽ�Ʋ����Ľ��ʼ�¼�����������", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�Ǳ��½��ʵĵ��ݣ����������
    gstrSQL = "select to_char(�շ�ʱ��,'yyyy-MM-dd') ����ʱ�� From ���˽��ʼ�¼ Where ID=" & lng����ID
    Call OpenRecordset(rsCheck, "ȡ��������")
    str�������� = Format(rsCheck!����ʱ��, "yyyyMM")
    str��ǰ���� = Format(zlDatabase.Currentdate, "yyyyMM")
    If str��ǰ���� <> str�������� Then
        MsgBox "ֻ�ܳ������µĽ��ʵ��ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '----׼����������----
    '��ȡҽ�����˵Ļ�����Ϣ
    gstrSQL = "Select ����,ҽ����,˳��� ����,��Ա���,���� From �����ʻ� Where ����=" & gintInsure & " And ����ID=" & lng����ID
    Call OpenRecordset(rsCheck, "��ȡҽ�����˵Ļ�����Ϣ")
    str���� = rsCheck!����
    strҽ���� = rsCheck!ҽ����
    str�����ı�� = rsCheck!����
    str��Ա��� = rsCheck!��Ա���
    str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21" _
                  , str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", True, "11")
    str���� = IIf(IsNull(rsCheck!����), "", rsCheck!����)
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDID", str����)                  ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", strҽ����)            ' ���˱���
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str�����ı��)        ' �����ı���
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str����)                ' ����
    Call InsertChild(mdomInput.documentElement, "PERSONTYPE", str��Ա���)          ' ��Ա���
    Call InsertChild(mdomInput.documentElement, "BILLNO", str����˳���)            ' ����˳���
    Call InsertChild(mdomInput.documentElement, "BALANCEID", str������)           ' ������
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)           ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) ' ��������
    
    '���ýӿ�
    If CommServer("RETHOSP") = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(rsTemp("����ID"), Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure("����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & rsTemp("����ID") & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & rsTemp("�������ý��") * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0," & rsTemp("�����Ը����") & "," & _
        cur�����ʻ� * -1 & ",'" & str������ & "',null,null,'" & str����˳��� & "')"
    Call ExecuteProcedure("����ҽ��")
    
    סԺ�������_���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Sub ��ѯǷ�ѵ�λ_����(ByVal str��λ���� As String)
'���ܣ����ýӿڲ�ѯǷ�ѵ�λ
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim str��ʾ As String
    
    If str��λ���� = "" Then Exit Sub
'    str��λ���� = String(12 - Len(str��λ����), "0") & str��λ����
    
    On Error GoTo errHandle
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", "0101")   ' �����ı���(����ҽ��)
    Call InsertChild(mdomInput.documentElement, "DEPTCODE", str��λ����)         ' ��λ����
    
    '���ýӿ�
    If CommServer("QUERYARREARDEPT") = False Then Exit Sub
    
    Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
    If nodRowset Is Nothing Then
        MsgBox "���˵�λ��Ƿ�������", vbInformation, gstrSysName
        Exit Sub
    End If
    '���ݱ���õ���������
    For Each nodRow In nodRowset.childNodes
        Select Case GetAttributeValue(nodRow, "INSUREKIND")
            Case "3"
                str��ʾ = str��ʾ & "������ҽ��"
            Case "8"
                str��ʾ = str��ʾ & "�����ҽ��"
        End Select
    Next
    
    If str��ʾ <> "" Then
        MsgBox "���˵�λ����������Ƿ�������" & Mid(str��ʾ, 2) & "��", vbInformation, gstrSysName
    Else
        MsgBox "���˵�λ��Ƿ�������", vbInformation, gstrSysName
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function ������Ϣ_����(ByVal lngErrCode As Long) As String
'���ܣ����ݴ���ŷ��ش�����Ϣ

End Function

Public Function ҽ����Ŀ_����(rsTemp As ADODB.Recordset) As Boolean
'���ܣ�ҽ������ҩƷĿ¼��ѯ
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim str���� As String, str���� As String, str����, strʧЧ As String
        
    On Error GoTo errHandle
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "ITEMCODE", "")         ' ҽ������
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", "0101") ' �����ı���(����ҽ��)
    
    '���ýӿ�
    If CommServer("QUERYSERVICE") = False Then Exit Function
    
    Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
    If nodRowset Is Nothing Then Exit Function
    For Each nodRow In nodRowset.childNodes
        str���� = GetAttributeValue(nodRow, "ITEMCODE")
        str���� = ToVarchar(Replace(GetAttributeValue(nodRow, "ITEMNAME"), "'", ""), 40)
        str���� = ToVarchar(zlCommFun.SpellCode(str����), 10)
        strʧЧ = GetAttributeValue(nodRow, "ISVALID")
        If str���� <> "" And strʧЧ <> "1" Then
            rsTemp.AddNew Array("CLASSCODE", "CODE", "NAME", "PY"), Array("1", str����, str����, str����)
            rsTemp.Update
        End If
    Next
    
    
    ҽ����Ŀ_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Public Function InitXML() As Boolean
'���ܣ���ʼ��XML�����������͸��ڵ�
    Dim pi As MSXML2.IXMLDOMProcessingInstruction
    Dim nodData As MSXML2.IXMLDOMElement
    
    On Error Resume Next
    
    Set mdomInput = New MSXML2.DOMDocument
    Set mdomOutput = New MSXML2.DOMDocument
    If Err <> 0 Then
        Err.Clear
        Exit Function
    End If
    
'    'XML����
'    Set pi = mdomInput.createProcessingInstruction("xml", "version=""1.0"" encoding=""GB2312"" standalone=""yes""")
'    mdomInput.appendChild pi
    
    '���ڵ�
    Set nodData = mdomInput.createElement("DATA")
    Set mdomInput.documentElement = nodData
    
    InitXML = True
End Function

Public Function InsertChild(nodParent As MSXML2.IXMLDOMElement, ByVal Name As String, ByVal Value As String) As MSXML2.IXMLDOMElement
'���ܣ���ָ��XMLԪ����������Ԫ��
    Set InsertChild = mdomInput.createElement(Name)
    InsertChild.Text = Value
    
    nodParent.appendChild InsertChild
End Function

Public Sub InsertAttrib(nodParent As MSXML2.IXMLDOMElement, ByVal Name As String, ByVal Value As String)
'���ܣ���ָ��XMLԪ������������
    Dim attTemp As MSXML2.IXMLDOMAttribute
    
    Set attTemp = mdomInput.createAttribute(Name)
    attTemp.Text = Value
    
    nodParent.setAttributeNode attTemp
End Sub
'
'Private Function CommServer(ByVal strFunction As String) As Boolean
''���ܣ���ҽ������������ͨѶ���õ�����ֵ
'    Dim cnComm As New ADODB.Connection
'    Dim rsTemp As New ADODB.Recordset
'    Dim lngID As Long
'
'    Dim lng���� As Long, lng��� As Long, strTemp As String, strInput As String
'    Dim timStart As Date, bln�Ѵ��� As Boolean
'
'
'    'Ϊ��ʵ��������ƣ���Ҫ��һ������
'    'ʹ�������Ŀ����Ϊ�˱�֤�������Ĵ�������ʹ�����������������
'    cnComm.ConnectionString = gcnOracle.ConnectionString
'    cnComm.Open
'
'    '�����Ĵ���
'    strInput = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?>" & vbCrLf & mdomInput.xml
'    lngID = zlDatabase.GetNextId("���սӿڱ�")
'    lng���� = Abs(Int(Len(strInput) / -2000))  '�п��ܴ�������󳤣�Ҫ�ֳɶ��в��ܱ���
'    On Error Resume Next
'    cnComm.BeginTrans
'    For lng��� = 1 To lng����
'        '�ֳ�������
'        strTemp = Replace(Mid(strInput, (lng��� - 1) * 2000 + 1, 2000), "'", "''")
'        gstrSQL = "insert into ���սӿڱ�(ID,���,����,����,�������,��������,״̬) values (" & _
'            lngID & "," & lng��� & "," & lng���� & ",'" & strFunction & "','" & strTemp & "','δ����',0)"
'        cnComm.Execute gstrSQL
'    Next
'    If Err <> 0 Then
'        '����
'        Err.Clear
'        cnComm.RollbackTrans
'        Exit Function
'    End If
'    cnComm.CommitTrans
'
'    On Error GoTo errHandle
'    '�ȴ���
'    timStart = Now
'    Do While True 'Ϊ�˱�֤ҽ���������Ĵ�����һ���ܵõ����գ���˲�ʹ�ó�ʱ�˳�     DateDiff("s", timStart, Now) < 600 'С��600����
'        DoEvents
'        If rsTemp.State = adStateOpen Then rsTemp.Close
'        gstrSQL = "select �������� from ���սӿڱ� where ID=" & lngID & " and ��������<>'δ����' order by ���"
'        rsTemp.Open gstrSQL, cnComm, adOpenStatic, adLockReadOnly
'
'        If rsTemp.EOF = False Then
'            'ȡ�÷���ֵ��
'            strTemp = ""
'            Do Until rsTemp.EOF
'                strTemp = strTemp & IIf(IsNull(rsTemp("��������")), "", rsTemp("��������"))
'                rsTemp.MoveNext
'            Loop
'
'            If mdomOutput.loadXML(strTemp) = False Then
'                MsgBox "ҽ������������ֵ��ʽ����ȷ��", vbInformation, gstrSysName
'            Else
'                '�ٶ����������Ƿ�ɹ����з���
'                If Val(GetElemnetValue("RETCODE")) = 0 Then
'                    '���óɹ�
'                    CommServer = True
'                Else
'                    '����ʧ��
'                    strTemp = GetElemnetValue("INFO")
'                    If strTemp = "" Then strTemp = "����������ʧ�ܡ�"
'                    MsgBox "ҽ�����������ش���" & vbCrLf & vbCrLf & strTemp, vbInformation, gstrSysName
'                End If
'            End If
'            bln�Ѵ��� = True
'            Exit Do
'        End If
'    Loop
'
'    If bln�Ѵ��� = False Then
'        MsgBox "��ҽ�����������ӳ�ʱ��", vbInformation, gstrSysName
'    End If
'errHandle:
'    cnComm.Execute "Delete from ���սӿڱ� where id=" & lngID '���۳ɹ���񣬶�������ɾ��
'End Function

Public Function CommServer(ByVal strFunction As String) As Boolean
'���ܣ�����ҽ������
    Dim objҽ�� As Object
    Dim InvokeServer As String '����ǰ�÷������ķ���ֵ
    Dim strInput As String, strServer As String
    
    On Error Resume Next
    '�����ȫ�ֱ�������ʱ����ʱ��Ⱥܾã�������Դ�����ԭ��
    strServer = Get���ղ���_����("ҽ��������")
    If strServer = "" Then
        Set objҽ�� = CreateObject("HospCOMSvr.HospCOMServer")
    Else
        Set objҽ�� = CreateObject("HospCOMSvr.HospCOMServer", strServer)
    End If
    If Err <> 0 Then
        MsgBox "�޷�����ҽ���ӿڲ�����HospCOMSvr.HospCOMServer����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�����Ĵ���
    strInput = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?>" & vbCrLf & mdomInput.xml
    
    Select Case strFunction
        Case "READCARD"         '���ʶ��/��Ϣ��ȡ
            InvokeServer = objҽ��.ReadCard("ZFRJ", strInput)
        Case "READCARD_M"       '���ʶ��/��Ϣ��ȡ���ֹ���ʽ��
            InvokeServer = objҽ��.ReadCard_M("ZFRJ", strInput)
        Case "MODIFYCARD"
            InvokeServer = objҽ��.MODIFYCARD("ZFRJ", strInput)
        'Modified By ���� 2004-05-25 ԭ��ҽ���ӿڱ䶯
        '------------------------------------------------
        Case "GETCLINNO"        '����Һ�
            InvokeServer = objҽ��.GETCLINNO("ZFRJ", strInput)
        '------------------------------------------------
        Case "CALCLIN"          '��ͨ����֧��
            InvokeServer = objҽ��.CALCLIN("ZFRJ", strInput)
        Case "CALSPECCLIN"      '��������֧��
            InvokeServer = objҽ��.CALSPECCLIN("ZFRJ", strInput)
        Case "RETCLIN"          '�շѳ���
            InvokeServer = objҽ��.RETCLIN("ZFRJ", strInput)
        Case "HOSPREG"          'סԺ�Ǽ�
            InvokeServer = objҽ��.HOSPREG("ZFRJ", strInput)
        Case "HOSPOUT"          '��Ժ�Ǽ�
            InvokeServer = objҽ��.HOSPOUT("ZFRJ", strInput)
        Case "CALHOSP"          'סԺ֧��
            InvokeServer = objҽ��.CALHOSP("ZFRJ", strInput)
        Case "RETHOSP"          '���ʳ���
            InvokeServer = objҽ��.RETHOSP("ZFRJ", strInput)
        Case "SETRECKONINGTYPE"
            InvokeServer = objҽ��.SETRECKONINGTYPE("ZFRJ", strInput)
        Case "QUERYHOSPSINGLEILLNESS"   '��������������
            InvokeServer = objҽ��.QUERYHOSPSINGLEILLNESS("ZFRJ", strInput)
        Case "QUERYSERVICE"     'ҽ������ҩƷĿ¼��ѯ
            InvokeServer = objҽ��.QUERYSERVICE("ZFRJ", strInput)
        Case "QUERYARREARDEPT"
            InvokeServer = objҽ��.QUERYARREARDEPT("ZFRJ", strInput)
        Case "GETHOSPSINGLEILLNESS"
            InvokeServer = objҽ��.GETHOSPSINGLEILLNESS("ZFRJ", strInput)
        Case Else
            MsgBox "����ҽ���ӿڷ����仯���޷�����ִ�н��ף���������ṩ����ϵ��", vbInformation, gstrSysName
            Exit Function
    End Select
    
    '�ϵ����ô�
    If InvokeServer = "" Then
        '����ʧ�ܣ����ع̶��Ĵ�����Ϣ
        InvokeServer = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?><DATA><RETCODE>-1</RETCODE><INFO>ҽ������������ʧ��</INFO></DATA>"
    End If
            
    If mdomOutput.loadXML(InvokeServer) = False Then
        MsgBox "ҽ������������ֵ��ʽ����ȷ��", vbInformation, gstrSysName
    Else
        '�ٶ����������Ƿ�ɹ����з���
        If Val(GetElemnetValue("RETCODE")) = 0 Then
            '���óɹ�
            CommServer = True
        Else
            '����ʧ��
            InvokeServer = GetElemnetValue("INFO")
            If InvokeServer = "" Then InvokeServer = "����������ʧ�ܡ�"
            MsgBox "ҽ�����������ش���" & vbCrLf & vbCrLf & InvokeServer, vbInformation, gstrSysName
        End If
    End If
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

Public Function SetElemnetValue(ByVal Name As String, ByVal Value As String) As Boolean
'���ܣ��õ�ָ��Ԫ�ص�ֵ
    Dim xmlElement As MSXML2.IXMLDOMElement
    
    Set xmlElement = mdomInput.documentElement.selectSingleNode(Name)
    If Not xmlElement Is Nothing Then
        '�ҵ�ָ����Ԫ��
        xmlElement.nodeTypedValue = Value
        SetElemnetValue = True
    End If
End Function

Public Function GetElemnetValue(ByVal Name As String) As String
'���ܣ��õ�ָ��Ԫ�ص�ֵ
    Dim xmlElement As MSXML2.IXMLDOMElement
    
    Set xmlElement = mdomOutput.documentElement.selectSingleNode(Name)
    If Not xmlElement Is Nothing Then
        '�ҵ�ָ����Ԫ��
        GetElemnetValue = xmlElement.Text
'    Else
'        'ȡ��
'        Debug.Assert False
    End If
End Function

Public Function GetAttributeValue(xmlElement As MSXML2.IXMLDOMElement, ByVal Name As String) As String
'���ܣ��õ�ָ�����Ե�ֵ
    Dim varAttribute As Variant
    
    varAttribute = xmlElement.getAttribute(Name)
    If IsNull(varAttribute) = False Then
        GetAttributeValue = varAttribute
    End If
End Function

Public Function Get��֤_����(str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, _
                ByVal lng����ID As Long, Optional blnǿ��ˢ�� As Boolean = False) As Boolean
'���ܣ��õ�ҽ�����˵Ļ������������֤��Ϣ
    Dim rsTemp As New ADODB.Recordset
    Dim strtemp As String
    
    If blnǿ��ˢ�� = False And lng����ID > 0 Then
        '�����ݿ��ж����Ѵ洢��ֵ
        gstrSQL = "select ����,ҽ����,˳���,���� from �����ʻ� where ����ID=" & lng����ID & " and ����=" & gintInsure
        Call OpenRecordset(rsTemp, "����ҽ��")
        
        If rsTemp.EOF = False Then
            strtemp = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            If strtemp = mstr���� And mstr���� <> "" Then
                '��ͬһ����
                str���� = mstr����
                str���� = mstr����
            Else
                str���� = strtemp
                str���� = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            End If
            
            strҽ���� = IIf(IsNull(rsTemp("ҽ����")), "", rsTemp("ҽ����"))
            str�����ı�� = IIf(IsNull(rsTemp("˳���")), "", rsTemp("˳���"))
            
            Get��֤_���� = True
            Exit Function
        End If
    End If
    
    If frmIdentify����.GetIdentify(TYPE_������, str����, strҽ����, str�����ı��, str����, True, True) = False Then
        Exit Function
    Else
        'ˢ����Ȼ��ȷ����Ҫ����Ƿ���ǵ�ǰ���˵�
            str���� = Split(str����, "^")(0)
            If lng����ID > 0 Then
            '�����ݿ��ж����Ѵ洢��ֵ
            gstrSQL = "select ����,ҽ����,˳��� from �����ʻ� where ����ID=" & lng����ID & " and ����=" & gintInsure
            Call OpenRecordset(rsTemp, "����ҽ��")
            
            If str���� <> IIf(IsNull(rsTemp("����")), "", rsTemp("����")) Or strҽ���� <> IIf(IsNull(rsTemp("ҽ����")), "", rsTemp("ҽ����")) Then
                MsgBox "��ǰʹ�õĿ��벡�˲�����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
    End If
    
    Get��֤_���� = True
End Function

Public Function OwnerUser(ByVal strUserName As String) As Boolean
    Dim RecUser As New ADODB.Recordset
    '�жϵ�ǰ�û��ǲ���������
    OwnerUser = True
    With RecUser
        If .State = 1 Then .Close
        .Open "Select Count(*) ������ From ZlSystems Where ������='" & strUserName & "'", gcnOracle
        
        If Not .EOF Then
            If Not IsNull(!������) Then
                If !������ = 0 Then OwnerUser = False
            End If
        End If
    End With
End Function

Public Function Subject(ByVal strData As String) As String
    Dim rsSubject As New ADODB.Recordset
    '���ض�Ӧ�Ĺ�����Ŀ����
    gstrSQL = "" & _
             " Select B.����,B.���,A.����ֵ ������Ŀ����   " & _
             " From ���ղ��� A,�շ���� B " & _
             " Where A.���>=6 And A.����=" & gintInsure & " And A.������=B.���� And B.����='" & strData & "'"
    Call OpenRecordset(rsSubject, "��ȡ��Ӧ�Ĺ�����Ŀ����")
    
    If rsSubject.EOF Then
        Subject = "11"  '�޶�Ӧ��Ŀ���ض�Ӧ�Ĺ�����Ŀ����'11',��ʾ����
    Else
        Subject = rsSubject!������Ŀ����
    End If
End Function

Public Function ����Һ�_����(ByVal lng����ID As Long, ByVal lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim datCurr As Date
    Dim str���㷽ʽ As String, arr���㷽ʽ
    Dim intTotal  As Integer, intStart As Integer
    Dim cur�ʻ���� As Double, cur�����ʻ� As Currency
    Dim curҽ������ As Currency, cur���ͳ�� As Currency
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, str����˳��� As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    gstrSQL = "Select B.����ID,B.����,B.ҽ����,B.���� " & _
        " From ����Ԥ����¼ A,�����ʻ� B " & _
        " Where A.����ID=B.����ID And B.����=" & gintInsure & _
        "       And A.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "ȡ����ID")
    If rsTemp.EOF Then Exit Function
    lng����ID = rsTemp!����ID
    If Get��֤_����(str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
    
    datCurr = zlDatabase.Currentdate()
    
    'ȡ�ʻ����
    gstrSQL = "Select Nvl(�ʻ����,0) �ʻ���� From �����ʻ� Where ����=" & gintInsure & " And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "ȡ�ʻ����")
    cur�ʻ���� = rsTemp!�ʻ����
    
    '��XML����ֵ
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", strҽ����)     ' ���˱���
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str�����ı��) ' �����ı���
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) ' ��������
    
    '���ýӿ�
    If CommServer("GETCLINNO") = False Then Exit Function
    str����˳��� = GetElemnetValue("BILLNO")
    
    gstrSQL = "Select ����ID,�շ�ϸĿID,����*NVL(����,1) AS ����,��׼���� AS ����,'  ' AS ժҪ" & _
        " From ���˷��ü�¼ " & _
        " Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Call OpenRecordset(rsTemp, "����ҽ��")
    If Not �����������_����(rsTemp, str���㷽ʽ) Then Exit Function
    
    '�ֽ���ֽ��㷽ʽ
    arr���㷽ʽ = Split(str���㷽ʽ, "|")
    intTotal = UBound(arr���㷽ʽ)
    For intStart = 0 To intTotal
        Select Case Split(arr���㷽ʽ(intStart), ";")(0)
        Case "�����ʻ�"
            cur�����ʻ� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        Case "ҽ������"
            curҽ������ = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        Case "���ͳ��"
            cur���ͳ�� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        End Select
    Next
    
    If Not �������_����(lng����ID, cur�����ʻ�, "") Then Exit Function
    
   '��Ҫ����������
    str���㷽ʽ = ""
    If cur�����ʻ� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & cur�����ʻ�
    If curҽ������ <> 0 Then str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & curҽ������
    If cur���ͳ�� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||���ͳ��|" & cur���ͳ��
    If str���㷽ʽ <> "" Then
        str���㷽ʽ = Mid(str���㷽ʽ, 3)
        gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',0)"
        Call ExecuteProcedure("����Ԥ����¼")
    End If
    
    ����Һ�_���� = True
    
    Call frm������Ϣ.ShowMe(lng����ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������㷽ʽ_����(ByVal lng����ID As Long, ByVal frmParent As Object) As Boolean
    �������㷽ʽ_���� = frm�������㷽ʽ.ShowSelect(lng����ID, TYPE_������, frmParent)
End Function

'������
'txtEdit(0).Text = "GY0001"
'txtEdit(1).Text = "01"
'txtEdit(2).Text = "01"
'str��� = "32"
'str���� = "����"
'str�Ա� = "��"
'str���֤���� = "510224770909071"
'str��Ա��� = "ʡ������"
'cur�ʻ���� = 500
'str����˳��� = "00000001"
'str������ = "JS000001"
'str���㷽ʽ = "�����ʻ�;1;0|ҽ������;2;0"
'curͳ��֧�� = 2

