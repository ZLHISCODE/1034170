Attribute VB_Name = "mdl����"
Option Explicit
#Const gblnTest = 0     '1-����

'�����޸�˵����
'�޸�ʱ�䣺2005-01-14
'�޸��ˣ�����
'�޸����ݣ����������ӿ�(SetBearingFlag��UploadICD)�������󲿷ֽӿ���������������
Public mdomInput As MSXML2.DOMDocument
Public mdomOutput As MSXML2.DOMDocument

Private mblnInit As Boolean
Private mstr���� As String
Private mstr���� As String
Private mstrҽ���� As String
Private mdbl��� As Double
Private mlng����ID As Long
Private mblnҽ����Ժ As Boolean         '��Ժʱ�Ƿ�ͬ������ҽ����Ժ
Private mbln��������� As Boolean
'����������
Private mint���㷽ʽ As Integer
Private mstr�����ֱ��� As String

Private objҽ�� As Object
Private obj���� As Object
Public Const mstrҽ�����ı���_���� As String = "0101"
Public gcnGYYB As New ADODB.Connection

Public Function ҽ����ʼ��_����() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    Dim strUser As String, strPass As String, strServer As String
    Dim rsTemp As New ADODB.Recordset
    
    If mblnInit Then
        ҽ����ʼ��_���� = True
        Exit Function
    End If
    
    On Error Resume Next
    Set mdomInput = New MSXML2.DOMDocument
    If Err <> 0 Then
        MsgBox "���ܴ���XML����������ע��msxml3.dll������", vbInformation, gstrSysName
        Exit Function
    End If
    
    Dim strYBServer As String
    On Error Resume Next
    #If gblnTest = 1 Then
        Set objҽ�� = CreateObject("GYSYB.CLSGYSYB")
        If Err <> 0 Then
            MsgBox "���ص��Բ���ʱ����������Ϣ���£�" & vbCrLf & Err.Description, vbInformation, gstrSysName
            Exit Function
        End If
        Set obj���� = objҽ��
    #Else
        '�����ȫ�ֱ�������ʱ����ʱ��Ⱥܾã�������Դ�����ԭ��
        strYBServer = Get���ղ���_����("ҽ��������")
        If strYBServer = "" Then
            Set objҽ�� = CreateObject("HospCOMSvr.HospCOMServer")
            Set obj���� = CreateObject("HospRecSvr.HospRecServer")
        Else
            Set objҽ�� = CreateObject("HospCOMSvr.HospCOMServer", strYBServer)
            Set obj���� = CreateObject("HospRecSvr.HospRecServer", strYBServer)
        End If
        If Err <> 0 Then
            MsgBox "�޷�����ҽ���ӿڲ�����HospCOMSvr.HospCOMServer����", vbInformation, gstrSysName
            Exit Function
        End If
    #End If
    
    'ȡ���ղ���
    On Error GoTo errHand
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & TYPE_������
    Call OpenRecordset(rsTemp, "ȡ���ղ���")
    Do While Not rsTemp.EOF
        If rsTemp!������ = "ҽ���û���" Then
            strUser = Nvl(rsTemp!����ֵ)
        ElseIf rsTemp!������ = "ҽ���û�����" Then
            strPass = Nvl(rsTemp!����ֵ)
        ElseIf rsTemp!������ = "ҽ��������1" Then
            strServer = Nvl(rsTemp!����ֵ)
        ElseIf rsTemp!������ = "��Ժ����" Then
            mblnҽ����Ժ = (Nvl(rsTemp!����ֵ, 0) = 0)
        End If
        rsTemp.MoveNext
    Loop
    If Not OraDataOpen(gcnGYYB, strServer, strUser, strPass, True) Then Exit Function
    
    mblnInit = True
    ҽ����ʼ��_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
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
    Dim STR���� As String, str�Ա� As String, str���֤���� As String, lng���� As Long
    Dim str�������� As String, str��Ա��� As String, str��λ���� As String, str��λ���� As String, strҽ���չ���Ա As String
    Dim strIdentify As String, str���� As String, lng����ID As Long
    Dim bln������־ As Boolean
    Dim rsTemp As New ADODB.Recordset, rs���� As ADODB.Recordset
    
    '��ʼ��һЩ�������ڳ�����;�˳�ʱֵȴ�Ѿ�����
    mstr���� = "": mstr���� = ""
    If frmIdentify����.GetIdentify(bytType, str����, strҽ����, str�����ı��, str����, bln������־) = False Then
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
    STR���� = GetElemnetValue("PERSONNAME")
    str�Ա� = GetElemnetValue("SEX")
    str�Ա� = Switch(str�Ա� = "1", "��", str�Ա� = "2", "Ů", str�Ա� = "9", "����", True, str�Ա�)
    str���֤���� = GetElemnetValue("PID")
    
    str�������� = AddDate(GetElemnetValue("BIRTHDAY"))
    If IsDate(str��������) = True Then
        lng���� = DateDiff("yyyy", CDate(str��������), zldatabase.Currentdate)
    Else
        str�������� = ""
    End If
    
    str��Ա��� = GetElemnetValue("PERSONTYPE")
    str��Ա��� = Switch(str��Ա��� = "11", "��ְ", str��Ա��� = "21", "����" _
                      , str��Ա��� = "32", "ʡ������", str��Ա��� = "34", "��������", True, "����")
    str��λ���� = ToVarchar(GetElemnetValue("DEPTCODE"), 12)
    str��λ���� = ToVarchar(GetElemnetValue("DEPTNAME"), 36) '�ֶγ��ȱ���50�������ڻ�Ҫ������뼰����
    cur�ʻ���� = Val(GetElemnetValue("ACCTBALANCE"))
    str������� = Val(GetElemnetValue("INSURETYPE"))
    strҽ���չ���Ա = Val(GetElemnetValue("CAREPSNFLAG"))
    
    '����;ҽ����;����;����;�Ա�;��������;���֤;������λ
    'ҽ���ŵ�һλΪ������
    '�ѷֺ��滻�ɶ���
    strIdentify = Replace(str����, ";", ",") & ";" & strҽ���� & ";" & str���� & ";" & STR���� & ";" & str�Ա� & ";" & str�������� & ";" & str���֤���� & ";" & str��λ���� & "(" & str��λ���� & ")"
    strIdentify = Replace(strIdentify, " ", "")
    
    '���������סԺ
    'Modified By ���� 2003-12-03 ������ ԭ����Ժʱȡ������ѡ�񣬸�Ϊ���������ʱ�����û�в��֣�����ѡ��
    If (bytType = id�����շ� And Get���ղ���_����("֧����������") = "1") Or bytType = id��Ժ�Ǽ� Then
        gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                " From ���ղ��� A where A.����=" & TYPE_������
        
        Set rs���� = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ������")
        If Not rs���� Is Nothing Then
            lng����ID = rs����("ID")
        End If
    End If
    
    str���� = ";"                                       '8.���Ĵ���
    str���� = str���� & ";"                             '9.˳���  ����ҽ�����ڱ���ҽ�������ı��루���⽨��ҽ�����ģ�
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
    
    lng����ID = BuildPatiInfo(bytType, strIdentify & str����, lng����ID, TYPE_������)
    '���ظ�ʽ:�м���벡��ID
    If lng����ID <> 0 Then
        ��ݱ�ʶ_���� = strIdentify & ";" & lng����ID & str����
        
        mstr���� = str����
        mstr���� = str����
        
        '���µ�ǰҽ�����˵ı�������Լ�ҽ���չ���Ա��־
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'�������','''" & str������� & "''')"
        Call zldatabase.ExecuteProcedure(gstrSQL, "���汣�����")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'ҽ���չ���Ա','''" & strҽ���չ���Ա & "''')"
        Call zldatabase.ExecuteProcedure(gstrSQL, "����ҽ���չ���Ա")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'������־','''" & IIf(bln������־, 1, 0) & "''')"
        Call zldatabase.ExecuteProcedure(gstrSQL, "����������־")
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
        gstrSQL = "Select �ʻ���� From �����ʻ� where ����=" & TYPE_������ & " and ����=0 and ҽ����='" & strSelfNo & "'"
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

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String, Optional strAdvance As String) As Boolean
'������rsDetail     ������ϸ(����)
'      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim str��Ŀ���� As String
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, str��Ա��� As String, str������� As String
    Dim dbl�����ʻ� As Double, dbl����Ա���� As Double, dblHIS�ܶ� As Double, dbl�����ܷ��� As Double, dbl��� As Double, dblTOTAL As Double
    Dim lng����ID As Long, str�������� As String, datCurr As Date
    Dim rsTemp As New ADODB.Recordset
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    On Error GoTo errHandle
    
    '���ӵ������Ժ󣬷��ز������ӣ������ܷ��ã�������Ҫ���Ӹ�����ֶΣ�������֤HIS����������ֽ�֧����ҽ��һ�£���ʽ���£�
    '���=HIS�ܷ���-�����ܷ��ã��ֽ�֧��=HIS�ܷ���-ͳ��֧��-���
    
    If rs��ϸ.RecordCount = 0 Then
        str���㷽ʽ = "�����ʻ�;0;0"
        �����������_���� = True
        Exit Function
    End If
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ("����ID")
    
    '�жϸò����Ƿ�����������
    gstrSQL = "select A.��Ա���,A.�������,B.���� from �����ʻ� A,���ղ��� B where A.����ID=" & lng����ID & " and A.����=" & TYPE_������ & "  and A.����ID=B.ID(+)"
    Call OpenRecordset(rsTemp, "����Ԥ��")
    If rsTemp.EOF = False Then
        str�������� = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        str��Ա��� = Nvl(rsTemp("��Ա���"), "")
        str������� = Nvl(rsTemp!�������)
        'ת����Ա���
        str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21" _
                      , str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", True, "11")
    End If
    datCurr = zldatabase.Currentdate
    
    mint���㷽ʽ = 0: mstr�����ֱ��� = ""
    If str�������� = "" Then
ReChoose:
        '��ͨ����Ҫ��ѡ����㷽ʽ�뵥���ֽ���Ŀ¼�����㷽ʽ;�����ֱ��룩
        mstr�����ֱ��� = ���ý��㷽ʽ_����(lng����ID, Nothing, False)
        If mstr�����ֱ��� = "" Then mstr�����ֱ��� = ";"
        mint���㷽ʽ = Val(Split(mstr�����ֱ���, ";")(0))
        mstr�����ֱ��� = Split(mstr�����ֱ���, ";")(1)
    End If
    
    If Get��֤_����(0, str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDDATA", str����)           ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "INSURETYPE", str�������)     ' �������
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str����)         ' ����
    If str�������� <> "" Then '��������
        '����8λ����
        str�������� = String(8 - Len(str��������), "0") & str��������
        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str��������)         '���ֲ�����
    End If
    Call InsertChild(mdomInput.documentElement, "ISCAL", 0)         ' �Ƿ����
    Call InsertChild(mdomInput.documentElement, "CALETYPE", mint���㷽ʽ)         ' ���㷽ʽ
    Call InsertChild(mdomInput.documentElement, "SINGLEILLNESSCODE", mstr�����ֱ���)         ' �����ֽ������
    Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", "0")     ' �˻�֧����
    Call InsertChild(mdomInput.documentElement, "INVOICENO", "") ' ��Ʊ��
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) ' ��������
    Call InsertChild(mdomInput.documentElement, "STARTDATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) '������ʼ����ʱ��
    Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' ������ϸ
    
    Do Until rs��ϸ.EOF
        gstrSQL = "SELECT C.ҩƷID,C.���,E.���� AS ����  FROM ҩƷĿ¼ C,ҩƷ��Ϣ D,ҩƷ���� E WHERE C.ҩƷID=" & rs��ϸ("�շ�ϸĿID") & " AND C.ҩ��ID=D.ҩ��ID AND D.����=E.����"
        gstrSQL = "select A.���,A.����,B.��Ŀ����,nvl(A.���,F.���) AS ���,F.����,A.���㵥λ from �շ�ϸĿ A,����֧����Ŀ B,(" & gstrSQL & _
                ") F where A.ID=" & rs��ϸ("�շ�ϸĿID") & " and A.ID=B.�շ�ϸĿID  AND A.Id=F.ҩƷID(+) and B.����=" & TYPE_������
        Call OpenRecordset(rsTemp, "����Ԥ��")
        If rsTemp.EOF = True Then
            MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
            Exit Function
        End If
        
        Set nodRow = InsertChild(nodRowset, "ROW", "")
        
        str��Ŀ���� = Nvl(rs��ϸ!���ձ���)
        '������ձ���Ϊ�գ�����Ҫ�û�ѡ�����
        If str��Ŀ���� = "" Then
            str��Ŀ���� = GetItemInsure_����(lng����ID, rs��ϸ!�շ�ϸĿID, True)
        End If
        If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsTemp!��Ŀ����)
        
        Call nodRow.setAttribute("ITEMCODE", ToVarchar(str��Ŀ����, 12))
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
        
        dblTOTAL = dblTOTAL + Round(rs��ϸ!ʵ�ս��, 2)
        rs��ϸ.MoveNext
    Loop
    
    '���ýӿ�
    If CommServer(IIf(str�������� <> "", "CALSPECCLIN", "CALCLIN")) = False Then Exit Function
    
    '��ͬ����Ⱥ�����ص�XML���ֶα�ʾ���岻ͬ������ֱ��ȡ������Ҫ�ֱ��ж�
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
    dbl����Ա���� = Val(GetElemnetValue("FUND3PAY"))
    str���㷽ʽ = str���㷽ʽ & "|ҽ�Ʋ���;" & dbl����Ա���� & ";0"
    
    dbl�����ܷ��� = Val(Format(GetElemnetValue("CALFEEALL"), "#0.00;-#0.00;0;"))
    dblHIS�ܶ� = Val(Format(GetElemnetValue("HOSPFEEALL"), "#0.00;-#0.00;0;"))
    
    '�ȱȽ��ܶ��Ƿ�һ��
    If Format(dblTOTAL, "#0.00") <> Format(dblHIS�ܶ�, "#0.00") Then
        MsgBox "HIS�ܶ���ҽ�����յ����ܷ��ò�һ�£���������㣡" & vbCrLf & _
            "HIS:" & Format(dblTOTAL, "#0.00") & Space(10) & "ҽ��:" & Format(dblHIS�ܶ�, "#0.00"), vbInformation, gstrSysName
        Exit Function
    End If
    
    dbl��� = dblHIS�ܶ� - dbl�����ܷ���
    If dbl��� <> 0 Then
        '���=HIS�ܷ���-�����ܷ��ã��ֽ�֧��=HIS�ܷ���-ͳ��֧��-���
        str���㷽ʽ = str���㷽ʽ & "|������;" & dbl��� & ";0"
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
    Dim str��Ŀ���� As String
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, str��Ա��� As String, str������� As String
    Dim strҽ�� As String, str���� As String, cur�������� As Double, curҽ���ܷ��� As Double, datCurr As Date
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
    
    gstrSQL = "Select A.ID,A.���,A.�շ�ϸĿID,A.��¼����,A.��¼״̬,A.����ID,A.NO,A.�Ǽ�ʱ��,A.������ as ҽ��,A.�Ǽ�ʱ��," & _
            "   A.����*A.���� as ����,A.��׼���� as ʵ�ʼ۸�,A.���ʽ��," & _
            "   A.�շ����,A.���ձ���,D.��Ŀ����,B.���� as ��Ŀ����,C.���� as ��������,nvl(B.���,F.���) AS ���,F.����,B.���㵥λ,A.ժҪ " & _
            " From (Select * From ���˷��ü�¼ Where ����ID=" & lng����ID & ") A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D " & _
            "     ,(SELECT C.ҩƷID,C.���,E.���� AS ����  FROM ���˷��ü�¼ A,ҩƷĿ¼ C,ҩƷ��Ϣ D,ҩƷ���� E WHERE A.����ID=" & lng����ID & " AND A.�շ�ϸĿID=C.ҩƷID AND C.ҩ��ID=D.ҩ��ID AND D.����=E.����) F " & _
            " Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID And A.�շ�ϸĿID=D.�շ�ϸĿID  AND A.ID=F.ҩƷID(+) And D.����=" & TYPE_������ & " And Nvl(A.���ӱ�־,0)<>9 And Nvl(A.��¼״̬,0)<>0" & _
            " Order by A.ID"
    Call OpenRecordset(rs��ϸ, "����ҽ��")
    If rs��ϸ.EOF = True Then
        MsgBox "û����д�շѼ�¼", vbExclamation, gstrSysName
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")
    strҽ�� = ToVarchar(IIf(IsNull(rs��ϸ("ҽ��")), UserInfo.����, rs��ϸ("ҽ��")), 20)
    str���� = ToVarchar(IIf(IsNull(rs��ϸ("��������")), UserInfo.����, rs��ϸ("��������")), 56)
    datCurr = zldatabase.Currentdate
    
    'һ��������ϸ����
    
    '�жϸò����Ƿ�����������
    gstrSQL = "select A.��Ա���,A.�������,B.���� from �����ʻ� A,���ղ��� B where A.����ID=" & lng����ID & " and A.����=" & TYPE_������ & "  and A.����ID=B.ID(+)"
    Call OpenRecordset(rsTemp, "����Ԥ��")
    If rsTemp.EOF = False Then
        str�������� = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        str��Ա��� = Nvl(rsTemp("��Ա���"), "")
        str������� = Nvl(rsTemp!�������)
        'ת����Ա���
        str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21" _
                      , str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", True, "11")
    End If
    
    If Get��֤_����(0, str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
        
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDDATA", str����)           ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "INSURETYPE", str�������)   ' �������
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str����)         ' ����
    If str�������� <> "" Then '��������
        '����8λ����
        str�������� = String(8 - Len(str��������), "0") & str��������
        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str��������)         '���ֲ�����
    End If
    Call InsertChild(mdomInput.documentElement, "ISCAL", 1)         ' �Ƿ����
    Call InsertChild(mdomInput.documentElement, "CALETYPE", mint���㷽ʽ)         ' ���㷽ʽ
    Call InsertChild(mdomInput.documentElement, "SINGLEILLNESSCODE", mstr�����ֱ���)         ' �����ֽ������
    Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", Format(cur�����ʻ�, "0.00")) ' �˻�֧����
    Call InsertChild(mdomInput.documentElement, "INVOICENO", "M" & "_" & rs��ϸ!��¼���� & "_" & rs��ϸ("NO")) ' ��Ʊ��
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(rs��ϸ("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")) ' ��������
    Call InsertChild(mdomInput.documentElement, "STARTDATE", Format(rs��ϸ("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")) '������ʼ����ʱ��
    Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' ������ϸ
    
    Do Until rs��ϸ.EOF
        cur�������� = cur�������� + rs��ϸ("���ʽ��")
        
        Set nodRow = InsertChild(nodRowset, "ROW", "")
        
        str��Ŀ���� = Nvl(rs��ϸ!���ձ���)
        '������ձ���Ϊ�գ�����Ҫ�û�ѡ�����
        If str��Ŀ���� = "" Then
            str��Ŀ���� = GetItemInsure_����(lng����ID, rs��ϸ!�շ�ϸĿID, False)
        End If
        If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rs��ϸ!��Ŀ����)
        
        Call nodRow.setAttribute("ITEMCODE", ToVarchar(str��Ŀ����, 12))
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
    Dim dblHIS�ܶ� As Double, dbl�����ܷ��� As Double, dbl��� As Double
    Dim cur����Ա�����𸶱�׼ As Double, cur����Ա�������� As Double, cur��ͨ���﹫��Ա�����ۼ� As Double, cur����Ա���� As Double, cur������޶��Ա���� As Double
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
    
    cur����Ա�����𸶱�׼ = Val(GetElemnetValue("STARTFEE2STD"))
    cur����Ա�������� = cur����
    cur��ͨ���﹫��Ա�����ۼ� = Val(GetElemnetValue("ENTERLMT3"))
    cur����Ա���� = Val(GetElemnetValue("FUND3PAY"))
    cur������޶��Ա���� = Val(GetElemnetValue("FUND3OVER"))
    curҽ���ܷ��� = Val(GetElemnetValue("FEEALL"))
    
    If str�������� = "" Then
        dbl�����ܷ��� = Val(GetElemnetValue("CALFEEALL"))
        dblHIS�ܶ� = Val(GetElemnetValue("HOSPFEEALL"))
        dbl��� = dblHIS�ܶ� - dbl�����ܷ���
    End If
    
    str������ = GetElemnetValue("BALANCEID")
    str����˳��� = GetElemnetValue("BILLNO")
    If str�������� <> "" Then
        str����˳��� = "����" & str�������� & str����˳��� '�Ѽ������������˳�������һ��
    Else
        str����˳��� = "��ͨ" & str����˳���         '��ʾ��ͨ����
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_������, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
                
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� + curͳ��֧�� + curͳ���Ը� + cur�����Ը� + cur�����Ը� + cur��ͳ�� + cur���Ը� & "," & _
        curͳ�ﱨ���ۼ� + curͳ��֧�� + cur��ͳ�� & "," & intסԺ�����ۼ� & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & ",0," & cur�����Ը� & "," & cur�������� & "," & _
        curȫ�Ը� & "," & cur�ҹ��Ը� & "," & curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & "," & cur���Ը� & "," & cur�����Ը� & "," & _
        cur�����ʻ� & ",'" & str������ & "',null,null,'" & str����˳��� & "',0,'" & AnalyseComputer & "','" & gstrVersion & "','" & IIf(str�������� <> "", "18", "11") & "','" & Mid(str����˳���, 3) & "'," & _
            "NULL,'" & str�������� & "','" & str������� & "',to_date('" & Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'))"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ﲻ�����������㷽ʽ���������㷽ʽ���㣨��ȷ��ȡֵ��Χ�Ǵ�1��ʼ��
    gstrSQL = "zl_���㸽����Ϣ_Insert (" & lng����ID & "," & cur����Ա�����𸶱�׼ & "," & cur����Ա�������� & "," & cur��ͨ���﹫��Ա�����ۼ� & "," & cur����Ա���� & "," & cur������޶��Ա���� & ",0,0," & _
        "'" & mstr�����ֱ��� & "'," & mint���㷽ʽ & ",NULL,0," & dbl�����ܷ��� & "," & curҽ���ܷ��� & ")"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
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
    Dim bln���� As Boolean
    Dim str֧������ As String
    
    On Error GoTo errHandle
    
    '�˷�
    '�ж��Ƿ��н��ʼ�¼�������˵����סԺ����ʵ�ֵ�
    gstrSQL = "Select 1 from ���˽��ʼ�¼ where ID=" & lng����ID
    Call OpenRecordset(rsTemp, "�ж��Ƿ��н��ʼ�¼�������˵����סԺ����ʵ�ֵ�")
    If rsTemp.RecordCount = 0 Then
        gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B " & _
                  " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Else
        gstrSQL = "select distinct A.ID AS ����ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
            " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=" & lng����ID
    End If
    Call OpenRecordset(rsTemp, "�����˷�")
    lng����ID = rsTemp("����ID")
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=" & TYPE_������ & " and ��¼ID=" & lng����ID
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
'    If Left(str����˳���, 2) = "����" Then
'        MsgBox "Ŀǰ��֧��������������ϡ�", vbInformation, gstrSysName
'        Exit Function
'    End If
    str֧������ = Nvl(rsTemp!ҽ�����)
    If str֧������ = "" Then str֧������ = IIf(Left(str����˳���, 2) = "����", "18", "11")
    str����˳��� = Mid(str����˳���, 3)
    curDate = zldatabase.Currentdate
    
    If Get��֤_����(0, str����, strҽ����, str�����ı��, str����, lng����ID, True) = False Then Exit Function
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "BILLNO", str����˳���)     ' ����˳���
    Call InsertChild(mdomInput.documentElement, "BALANCEID", str������)    ' ������
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", str֧������)   ' ֧�����
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)    ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(curDate, "yyyy-MM-dd HH:mm:ss"))  ' ��������
    
    '���ýӿ�
    bln���� = IS����(lng����ID)
    If CommServer("RETBALANCE", IIf(bln����, 1, 0)) = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_������, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & rsTemp("�������ý��") * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0," & rsTemp("�����Ը����") & "," & _
        cur�����ʻ� * -1 & ",'" & str������ & "',null,null,'" & Nvl(rsTemp("��ע")) & "'," & _
        "0,'" & AnalyseComputer & "','" & gstrVersion & "','" & str֧������ & "','" & Nvl(rsTemp!������ˮ��) & "'," & _
        "NULL,'" & Nvl(rsTemp!��������) & "','" & Nvl(rsTemp!����֢) & "',to_date('" & Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'))"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gstrSQL = "Select * From ���㸽����Ϣ Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ���㸽�Ӽ�¼", gstrSQL, gcnGYYB)
    If rsTemp.RecordCount <> 0 Then
        gstrSQL = "zl_���㸽����Ϣ_Insert (" & lng����ID & "," & -1 * Nvl(rsTemp!����Ա�����𸶱�׼, 0) & "," & -1 * Nvl(rsTemp!����Ա��������, 0) & "," & -1 * Nvl(rsTemp!��ͨ���﹫��Ա�����ۼ�, 0) & "," _
            & -1 * Nvl(rsTemp!����Ա����, 0) & "," & -1 * Nvl(rsTemp!������Ա����, 0) & ",0,0,'" & Nvl(rsTemp!�����ֱ���_����) & "'," & Nvl(rsTemp!���㷽ʽ, 0) & ",'" & Nvl(rsTemp!������) & "'," & _
            Nvl(rsTemp!���㷽ʽ, 0) & "," & -1 * Nvl(rsTemp!�����ܷ���, 0) & "," & -1 * Nvl(rsTemp!ҽ���ܷ���, 0) & ")"
        gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
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
    Dim strTemp As String, str��ʾ As String, str��� As String, lng�α�ǰ��Ժ As Long
    Dim str֧������ As String
    On Error GoTo errHandle
    
    If Get��֤_����(1, str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
    
    '�жϸò����Ƿ�α�ǰ��Ժ
    lng�α�ǰ��Ժ = 0
    If Get���ղ���_����("��Ժʱѡ��α�ǰ��Ժ") = "1" Then
        If MsgBox("�ò��˲α�ǰ�Ƿ��Ѿ���Ժ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            lng�α�ǰ��Ժ = 1
        End If
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
    datCurr = zldatabase.Currentdate
    gstrSQL = "select A.��Ժ��ʽ,nvl(A.����Ժת��,0) as ����Ժת��,A.����ҽʦ,A.��Ժ����,A.��Ժ����,B.���� as ��Ժ����,C.סԺ��,D.�������,D.������־ " & _
        " from ������ҳ A,���ű� B,������Ϣ C,�����ʻ� D " & _
        " Where A.����ID=D.����ID And D.����=" & TYPE_������ & " And A.����ID=C.����ID and A.��Ժ����ID = B.ID And A.����ID =" & lng����ID & " And A.��ҳID = " & lng��ҳID
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    str֧������ = IIf(rsTemp("��Ժ��ʽ") = "ת��", "37", IIf(rsTemp("��Ժ��ʽ") = "�ƻ�����", "32", "31"))  ' ֧����� 31��סԺ��37��תԺ
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDDATA", str����)           ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str����)         ' ����
    Call InsertChild(mdomInput.documentElement, "INSURETYPE", Nvl(rsTemp!�������))   ' �������
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", str֧������)
    Call InsertChild(mdomInput.documentElement, "HOSPNO", ToVarchar(rsTemp("סԺ��"), 20))     ' סԺ��
    Call InsertChild(mdomInput.documentElement, "ISINHOSP", lng�α�ǰ��Ժ)     ' �α�ǰ����Ժ 1���ǣ�0����
    Call InsertChild(mdomInput.documentElement, "DIAGNOSES", str���) ' ���
    Call InsertChild(mdomInput.documentElement, "DOCTOR", ToVarchar(rsTemp("����ҽʦ"), 20)) ' ���ҽ��
    Call InsertChild(mdomInput.documentElement, "OFFICE", ToVarchar(rsTemp("��Ժ����"), 20)) ' ����
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
        strTemp = Left(str������Ϣ, 2)
        str������Ϣ = Mid(str������Ϣ, 41)
        
        str��ʾ = str��ʾ & Switch(strTemp = "11", "�����������", strTemp = "21", "��������", strTemp = "31", "������ͳ��Ƿ��", _
                                   strTemp = "32", "�����ͳ��δ�ɷ�", strTemp = "41", "��ͣ��", strTemp = "51", "���˱�")
        
    Loop
    str��ʾ = str��ʾ & GetElemnetValue("NOTE")
    If str��ʾ <> "" Then
        MsgBox "��ע���ҽ�����������" & Mid(str��ʾ, 2) & "��", vbInformation, gstrSysName
    End If
    
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(datCurr) & "," & _
        "0,0,0," & curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�������� & "," & cur�����ۼ� & _
         "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & "," & cur���ͳ���ۼ� & ",'" & ToVarchar(str��ʾ, 100) & "')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'˳���','''" & GetElemnetValue("BILLNO") & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�˳���")
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'�����ֱ���_����','''" & "" & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "��������ֱ���_����")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'���㷽ʽ','''" & "" & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "������㷽ʽ")
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, Optional ByVal bln��Ժ As Boolean = False) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
            'ȡ��Ժ�Ǽ���֤�����ص�˳���
            
    '�޸�˵��
    'ʱ�䣺2005-01-14
    '�޸��ˣ�����
    '�޸����ݣ���Ժ�Ǽǽӿ�������Ρ�ICD���룬Ҳ�����ṩ���ϴ�ICD����Ľӿڣ����뿭��ϵ���ݶ����ӿڲ��ϴ�ICD���룬���ϴ�ICD�������
    
    Dim strҽ���� As String
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str��� As String, str������� As String
    Dim str������ As String, str��Ժת�� As String, lngPos As Long
    
    On Error GoTo errHandle
    
    If mblnҽ����Ժ Or bln��Ժ Then
        '�����ݿ��ж����Ѵ洢��ֵ
        gstrSQL = "select ����,ҽ����,˳��� from �����ʻ� where ����ID=" & lng����ID & " and ����=" & TYPE_������
        Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
        
        strҽ���� = IIf(IsNull(rsTemp("ҽ����")), "", rsTemp("ҽ����"))
        
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
        datCurr = zldatabase.Currentdate
        gstrSQL = "select A.סԺҽʦ,A.��Ժ����,A.��Ժ����,B.���� as ��Ժ���� from ������ҳ A,���ű� B " & _
                 " Where A.��Ժ����ID = B.ID And A.����ID =" & lng����ID & " And A.��ҳID = " & lng��ҳID
        Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
        
        '��XML DomDocument������г�ʼ��
        If InitXML = False Then Exit Function
        Call InsertChild(mdomInput.documentElement, "PERSONCODE", strҽ����)     ' ���˱���
        Call InsertChild(mdomInput.documentElement, "DOCNO", str������)          ' ������
        Call InsertChild(mdomInput.documentElement, "DIAGNOSES", ToVarchar(str���, 128))          ' ���
        Call InsertChild(mdomInput.documentElement, "OTHERDIAGNOSES", ToVarchar(str�������, 128)) ' �������
        Call InsertChild(mdomInput.documentElement, "OUTTYPE", str��Ժת��)                        ' ת�����
        Call InsertChild(mdomInput.documentElement, "ICD", "")                       ' ICD��������
        Call InsertChild(mdomInput.documentElement, "DOCTOR", ToVarchar(rsTemp("סԺҽʦ"), 20))   ' ���ҽ��
        Call InsertChild(mdomInput.documentElement, "OFFICE", ToVarchar(rsTemp("��Ժ����"), 20))   ' ����
        Call InsertChild(mdomInput.documentElement, "REGDATE", Format(rsTemp("��Ժ����"), "yyyy-MM-dd HH:mm:ss")) ' ��Ժ����
        Call InsertChild(mdomInput.documentElement, "OPERATOR", UserInfo.����) ' ����Ա
        Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss"))  ' ��������
        
        '���ýӿ�
        If CommServer("HOSPOUT") = False Then Exit Function
    End If
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    MsgBox IIf(mblnҽ����Ժ, "�ɹ�����HIS��ҽ����Ժ��", "��������HIS��Ժ��"), vbInformation, gstrSysName
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim str˳��� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand

    gstrSQL = " Select ˳��� From �����ʻ� Where ����=" & TYPE_������ & " And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ������ˮ��")
    If rsTemp.RecordCount = 0 Then
        MsgBox "û���ҵ��ò��˵�ҽ��������", vbInformation, gstrSysName
        Exit Function
    End If
    str˳��� = Nvl(rsTemp!˳���)

    '�˴�����ҽ�����óɹ������м�飬���Ժʱ���ܽ�������HIS��Ժ
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "BILLNO", str˳���) ' ��Ժʱ��
    Call InsertChild(mdomInput.documentElement, "OPERATOR", UserInfo.����) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zldatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss"))  ' ��������
    Call CommServer("RETHOSPOUT")

    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ҽ��")

    ��Ժ�Ǽǳ���_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����(rsExse As Recordset, ByVal lng����ID As Long) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim cn�ϴ� As New ADODB.Connection, rsTemp As New ADODB.Recordset, rs���� As ADODB.Recordset
    Dim lng����ID As Long, str�������� As String, str��Ŀ���� As String
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, str��Ա��� As String
    Dim cur�����ʻ� As Double, curͳ��֧�� As Double, cur��ͳ�� As Double, cur�������� As Double
    Dim dbl�����ܷ���  As Double, dblHIS�ܷ��� As Double, dbl��� As Double
    Dim cur����Ա���� As Double, curҽ���չ˹���Ա���� As Double, cur������Ա���� As Double, int������־ As Integer
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
    gstrSQL = "select A.��Ա���,B.���� from �����ʻ� A,���ղ��� B where A.����ID=" & lng����ID & " and A.����=" & TYPE_������ & "  and A.����ID=B.ID(+)"
    Call OpenRecordset(rsTemp, "סԺԤ��")
    If rsTemp.EOF = False Then
        str�������� = Nvl(rsTemp!����)
        str��Ա��� = Nvl(rsTemp("��Ա���"), "")
        'ת����Ա���
        str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21" _
                      , str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", True, "11")
    End If
    
    mstr���� = ""
    mstr���� = ""
    If Get��֤_����(1, str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
    Screen.MousePointer = vbHourglass
    
    mbln��������� = False
    If MsgBox("�Ƿ������������㣨�������ڲ����ӷѵ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mbln��������� = True
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    'סԺ�������ֻҪ������˱��룬��ʽ����ʱ��Ҫ����ſ����ݼ�����
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", strҽ����)     ' ���˱���
    Call InsertChild(mdomInput.documentElement, "ISCAL", 0)         ' �Ƿ����
    Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", "0")     ' �˻�֧����
    Call InsertChild(mdomInput.documentElement, "INVOICENO", "") ' ��Ʊ��
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) ' ��������
    Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' ������ϸ
    
    rsExse.Sort = " �Ǽ�ʱ�� asc"
    Do Until rsExse.EOF
        If IIf(IsNull(rsExse("�Ƿ��ϴ�")), "0", rsExse("�Ƿ��ϴ�")) = "0" Then
            gstrSQL = "SELECT C.ҩƷID,C.���,E.���� AS ����  FROM ҩƷĿ¼ C,ҩƷ��Ϣ D,ҩƷ���� E WHERE C.ҩƷID=" & rsExse("�շ�ϸĿID") & " AND C.ҩ��ID=D.ҩ��ID AND D.����=E.����"
            gstrSQL = "select A.���,A.����,B.��Ŀ����,nvl(A.���,F.���) AS ���,F.����,A.���㵥λ from �շ�ϸĿ A,����֧����Ŀ B,(" & gstrSQL & _
                    ") F where A.ID=" & rsExse("�շ�ϸĿID") & " and A.ID=B.�շ�ϸĿID  AND A.Id=F.ҩƷID(+) and B.����=" & TYPE_������
            Call OpenRecordset(rsTemp, "סԺԤ��")
            If rsTemp.EOF = True Then
                MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
                Exit Function
            End If
            
            'ֻ�ϴ�ֻ���ݹ�������
            strҽ�� = ToVarchar(IIf(IsNull(rsExse("ҽ��")), UserInfo.����, rsExse("ҽ��")), 20)
            str���� = ToVarchar(IIf(IsNull(rsExse("��������")), UserInfo.����, rsExse("��������")), 56)
        
            str��Ŀ���� = Nvl(rsExse!���ձ���)
            '������ձ���Ϊ�գ�����Ҫ�û�ѡ�����
            If str��Ŀ���� = "" Then
                str��Ŀ���� = GetItemInsure_����(lng����ID, rsExse!�շ�ϸĿID, False)
            End If
            If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsExse!ҽ����Ŀ����)
            
            Set nodRow = InsertChild(nodRowset, "ROW", "")
            Call nodRow.setAttribute("ITEMCODE", ToVarchar(str��Ŀ����, 12))
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
        cur�������� = cur�������� + Round(rsExse("���"), 2)
        rsExse.MoveNext
    Loop
    
    '���ýӿ�
    If CommServer("CALHOSP", IIf(mbln���������, "1", "0")) = False Then Exit Function
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
    
'    <FUND3PAY>����Ա����֧��</FUND3PAY>
'    <CAREPAY>ҽ���չ���Ա�����Ա����</CAREPAY>
'    <FUND3OVER>������޶��Ա����</ FUND3OVER >
'    <BEARINGFLAG>������־</BEARINGFLAG>
    cur����Ա���� = Val(GetElemnetValue("FUND3PAY"))
    dbl�����ܷ��� = Val(Format(GetElemnetValue("CALFEEALL"), "#0.00;-#0.00;0;"))
    dblHIS�ܷ��� = Val(Format(GetElemnetValue("HOSPFEEALL"), "#0.00;-#0.00;0;"))
    dbl��� = dblHIS�ܷ��� - dbl�����ܷ���
    
    '�ȱȽ��ܶ��Ƿ�һ��
    If Format(cur��������, "#0.00") <> Format(dblHIS�ܷ���, "#0.00") Then
        MsgBox "HIS�ܶ���ҽ�����յ����ܷ��ò�һ�£���������㣡" & vbCrLf & _
            "HIS:" & Format(cur��������, "#0.00") & Space(10) & "ҽ��:" & Format(dblHIS�ܷ���, "#0.00"), vbInformation, gstrSysName
        Exit Function
    End If
    
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
    סԺ�������_���� = סԺ�������_���� & "|��ͳ��;" & cur��ͳ�� & ";0"
    סԺ�������_���� = סԺ�������_���� & "|ҽ�Ʋ���;" & Format(cur����Ա����, "#0.00;-#0.00;0;") & ";0"
    סԺ�������_���� = סԺ�������_���� & "|������;" & dbl��� & ";0"
    
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
    Dim lng��ҳID As Long
    Dim curȫ�Ը� As Double, cur�ҹ��Ը� As Double, curͳ��֧�� As Double
    Dim curͳ���Ը� As Double, cur�����Ը� As Double, cur�����Ը� As Double
    Dim cur��ͳ�� As Double, cur���Ը� As Double, cur�����ʻ� As Double, cur���� As Currency
    Dim cur����Ա���� As Double, curҽ���չ˹���Ա���� As Double, cur������Ա���� As Double, int������־ As Integer
    Dim dblHIS�ܷ��� As Double, dbl�����ܷ��� As Double, dbl��� As Double, dblҽ���ܷ��� As Double
    
    Dim int���㷽ʽ As Integer, str�����ֱ��� As String
    Dim int���㷽ʽ As Integer, str������ As String
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, datCurr As Date, strNO As String
    Dim str����˳��� As String, str������ As String
    Dim str֧������ As String, str������� As String
    
    If mlng����ID <> lng����ID Then
        MsgBox "�ò���û�����ҽ����Ԥ������������ܽ��н��㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHandle
    'ȡ��ҳID
    gstrSQL = " Select A.�������,B.סԺ���� AS ��ҳID,A.������,A.���㷽ʽ,A.�����ֱ���_����,A.���㷽ʽ,C.��Ժ��ʽ " & _
              " From �����ʻ� A,������Ϣ B,������ҳ C " & _
              " Where A.����ID=B.����ID And B.����ID=C.����ID And B.סԺ����=C.��ҳID And A.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "ȡ��ҳID")
    lng��ҳID = rsTemp!��ҳID
    str������� = Nvl(rsTemp!�������)
    str������ = Nvl(rsTemp!������)
    int���㷽ʽ = Nvl(rsTemp!���㷽ʽ, 1)
    str�����ֱ��� = Nvl(rsTemp!�����ֱ���_����)
    int���㷽ʽ = Nvl(rsTemp!���㷽ʽ, 0)
    str֧������ = IIf(rsTemp("��Ժ��ʽ") = "ת��", "37", IIf(rsTemp("��Ժ��ʽ") = "�ƻ�����", "32", "31"))      ' ֧����� 31��סԺ��37��תԺ
    
    '������ʻ�֧�����
    gstrSQL = "Select Nvl(��Ԥ��,0) as ��� From ����Ԥ����¼ Where ���㷽ʽ='�����ʻ�' And ��¼���� Not In (11,1) And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "סԺ����")
    If Not rsTemp.EOF Then cur�����ʻ� = rsTemp("���")
    '�󵥾ݺ�
    gstrSQL = "Select NO,�շ�ʱ�� From ���˽��ʼ�¼ Where ID=" & lng����ID
    Call OpenRecordset(rsTemp, "סԺ����")
    
    'XML�ĵ��Ѿ���ɳ�ʼ������ʱֻ��Ҫ���²���ֵ
    Call SetElemnetValue("ISCAL", "1")
    Call SetElemnetValue("ACCTWANTTOPAY", Format(cur�����ʻ�, "0.00"))
    Call SetElemnetValue("INVOICENO", "Z_" & rsTemp("NO"))
    Call SetElemnetValue("DODATE", Format(rsTemp("�շ�ʱ��"), "yyyy-MM-dd HH:mm:ss"))
    'Ԥ��ʱ�Ѿ����ݣ����ʲ���Ҫ�ٴ�����ϸ����
    Call SetElemnetValue("ROWSET", "")
    '���ýӿ�
    If CommServer("CALHOSP", IIf(mbln���������, "1", "0")) = False Then Exit Function
    
    curȫ�Ը� = Val(GetElemnetValue("FEEOUT"))
    cur�ҹ��Ը� = Val(GetElemnetValue("FEESELF"))
    cur���� = Val(GetElemnetValue("STARTFEE"))
    cur�����Ը� = Val(GetElemnetValue("ENTERSTARTFEE"))
    curͳ��֧�� = Val(GetElemnetValue("FUND1PAY")) + Val(GetElemnetValue("ALLOWFUND"))
    curͳ���Ը� = Val(GetElemnetValue("FUND1SELF"))
    cur��ͳ�� = Val(GetElemnetValue("FUND2PAY"))
    cur���Ը� = Val(GetElemnetValue("FUND2SELF"))
    cur�����Ը� = Val(GetElemnetValue("FEEOVER"))
    
'    <FUND3PAY>����Ա����֧��</FUND3PAY>
'    <CAREPAY>ҽ���չ���Ա�����Ա����</CAREPAY>
'    <FUND3OVER>������޶��Ա����</ FUND3OVER >
'    <BEARINGFLAG>������־</BEARINGFLAG>
    dblҽ���ܷ��� = Val(GetElemnetValue("FEEALL"))
    cur����Ա���� = Val(GetElemnetValue("FUND3PAY"))
    dbl�����ܷ��� = Val(GetElemnetValue("CALFEEALL"))
    dblHIS�ܷ��� = Val(GetElemnetValue("HOSPFEEALL"))
    dbl��� = dblHIS�ܷ��� - dbl�����ܷ���
    
    str������ = GetElemnetValue("BALANCEID")
    str����˳��� = GetElemnetValue("BILLNO")
    
    '��д�����
    datCurr = zldatabase.Currentdate
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_������, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
            
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� + curͳ��֧�� + curͳ���Ը� + cur�����Ը� + cur�����Ը� + cur��ͳ�� + cur���Ը� & "," & _
        curͳ�ﱨ���ۼ� + curͳ��֧�� + cur��ͳ�� & "," & intסԺ�����ۼ� & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & ",NULL," & cur�����Ը� & "," & _
        g��������.�������ý�� & "," & curȫ�Ը� & "," & cur�ҹ��Ը� & "," & _
        curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & "," & cur���Ը� & "," & cur�����Ը� & "," & cur�����ʻ� & "," & _
        "'" & str������ & "'," & lng��ҳID & ",null,'" & str����˳��� & "',0,'" & AnalyseComputer & "','" & gstrVersion & "','" & str֧������ & "','" & str����˳��� & "'," & _
            "NULL,'" & str�����ֱ��� & "','" & str������� & "',to_date('" & Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'))"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gstrSQL = "zl_���㸽����Ϣ_Insert (" & lng����ID & ",0,0,0," & cur����Ա���� & "," & cur������Ա���� & "," & curҽ���չ˹���Ա���� & "," & int������־ & "," & _
        "'" & str�����ֱ��� & "'," & int���㷽ʽ & ",'" & str������ & "'," & int���㷽ʽ & "," & dbl�����ܷ��� & "," & dblҽ���ܷ��� & ")"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    '���ս������
    gstrSQL = "zl_���ս������_insert(" & lng����ID & ",0," & curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & ",NULL)"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '������㷽ʽ���ǰ����嵥����Ա�����������Ա�����ҵ�ǰ��Ժ��ҽ�����ˣ�����ʾ����ԱΪ�ò��˰����Ժ����
    gstrSQL = "Select ������,��Ա��� From �����ʻ� Where ����=" & TYPE_������ & " And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ���㷽ʽ")
    If Right(rsTemp!������, 1) <> 4 And Not (rsTemp!��Ա��� = "��������" Or rsTemp!��Ա��� = "ʡ������") And ҽ�������Ѿ���Ժ(lng����ID) = False Then
        MsgBox "��Ϊ�òα���Ա�����Ժ������", vbInformation, gstrSysName
    End If
    
    סԺ����_���� = True
    
    '����ҽ����Ժ���������������HIS��Ժͬʱ����ҽ����Ժ�Ļ�������Ҫ�ڽ���ɹ������ҽ����Ժ���������ʧ�ܣ����Ա����ʻ����ٴΰ���ҽ����Ժ��
    If mblnҽ����Ժ = False And ҽ�������Ѿ���Ժ(lng����ID) Then
        Call ��Ժ�Ǽ�_����(lng����ID, lng��ҳID, True)
    End If
    
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
    Dim bln���� As Boolean
    Dim str֧������ As String
    
    On Error GoTo errHand
    curDate = zldatabase.Currentdate
    
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
        " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=" & lng����ID
    Call OpenRecordset(rsTemp, "�������")
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    'Ϊ�˽���ʱд���Ľ����������ٴη��ʼ�¼
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=" & TYPE_������ & " and ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, "�������")
    If rsTemp.EOF = True Then
        MsgBox "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    lng����ID = rsTemp!����ID
    str������ = IIf(IsNull(rsTemp!֧��˳���), "", rsTemp!֧��˳���)
    str����˳��� = IIf(IsNull(rsTemp!��ע), "", rsTemp!��ע)
    str֧������ = Nvl(rsTemp!ҽ�����)
    If str֧������ = "" Then
        '�Ӳ�������ȡ
        gstrSQL = " Select B.��Ժ��ʽ From ������Ϣ A,������ҳ B" & _
                  " Where A.����ID=B.����ID And A.סԺ����=B.��ҳID And A.����ID=" & lng����ID
        Call OpenRecordset(rsCheck, "ȡ��Ժ��ʽ")
        str֧������ = IIf(rsCheck("��Ժ��ʽ") = "ת��", "37", IIf(rsCheck("��Ժ��ʽ") = "�ƻ�����", "32", "31"))      ' ֧����� 31��סԺ��37��תԺ
    End If
'
'    '�ж��Ƿ�Ϊ������Ա
'    gstrSQL = "Select ��Ա��� From �����ʻ� Where ����ID=" & lng����ID & " And ����=" & TYPE_������
'    Call OpenRecordset(rsCheck, "�ж��Ƿ�Ϊ������Ա")
'    If Not (rsCheck!��Ա��� = "ʡ������" Or rsCheck!��Ա��� = "��������") Then
'        MsgBox "����ҽ�Ʋ����Ľ��ʼ�¼�����������", vbInformation, gstrSysName
'        Exit Function
'    End If
    
    '�Ǳ��½��ʵĵ��ݣ����������
    gstrSQL = "select to_char(�շ�ʱ��,'yyyy-MM-dd') ����ʱ�� From ���˽��ʼ�¼ Where ID=" & lng����ID
    Call OpenRecordset(rsCheck, "ȡ��������")
    str�������� = Format(rsCheck!����ʱ��, "yyyyMM")
    str��ǰ���� = Format(zldatabase.Currentdate, "yyyyMM")
    If str��ǰ���� <> str�������� Then
        MsgBox "ֻ�ܳ������µĽ��ʵ��ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '----׼����������----
    '��ȡҽ�����˵Ļ�����Ϣ
    gstrSQL = "Select ����,ҽ����,˳��� ����,��Ա���,���� From �����ʻ� Where ����=" & TYPE_������ & " And ����ID=" & lng����ID
    Call OpenRecordset(rsCheck, "��ȡҽ�����˵Ļ�����Ϣ")
    str���� = rsCheck!����
    strҽ���� = rsCheck!ҽ����
    str��Ա��� = rsCheck!��Ա���
    str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21" _
                  , str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", True, "11")
    str���� = IIf(IsNull(rsCheck!����), "", rsCheck!����)
    bln���� = (str��Ա��� = "32" Or str��Ա��� = "34")
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "BILLNO", str����˳���)            ' ����˳���
    Call InsertChild(mdomInput.documentElement, "BALANCEID", str������)           ' ������
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", str֧������)           ' ֧������
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)           ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) ' ��������
    
    '���ýӿ�
    If CommServer("RETBALANCE", IIf(bln����, 1, 0)) = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_������, rsTemp("����ID"), Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_������ & "," & rsTemp("����ID") & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & rsTemp("�������ý��") * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0," & rsTemp("�����Ը����") & "," & _
        -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & ",'" & Nvl(rsTemp!֧��˳���) & "',null,null,'" & Nvl(rsTemp!��ע) & "'," & _
        "0,'" & AnalyseComputer & "','" & gstrVersion & "','" & str֧������ & "','" & Nvl(rsTemp!������ˮ��) & "'," & _
        "NULL,'" & Nvl(rsTemp!��������) & "','" & Nvl(rsTemp!����֢) & "',to_date('" & Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'))"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gstrSQL = "Select * From ���㸽����Ϣ Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ���㸽�Ӽ�¼", gstrSQL, gcnGYYB)
    If rsTemp.RecordCount <> 0 Then
        gstrSQL = "zl_���㸽����Ϣ_Insert (" & lng����ID & "," & -1 * Nvl(rsTemp!����Ա�����𸶱�׼, 0) & "," & -1 * Nvl(rsTemp!����Ա��������, 0) & "," & -1 * Nvl(rsTemp!��ͨ���﹫��Ա�����ۼ�, 0) & "," _
            & -1 * Nvl(rsTemp!����Ա����, 0) & "," & -1 * Nvl(rsTemp!������Ա����, 0) & "," & -1 * Nvl(rsTemp!ҽ���չ���Ա�����Ա����, 0) & "," & rsTemp!������־ & "," & _
            "'" & Nvl(rsTemp!�����ֱ���_����) & "'," & Nvl(rsTemp!���㷽ʽ, 0) & ",'" & Nvl(rsTemp!������) & "'," & Nvl(rsTemp!���㷽ʽ, 1) & "," & -1 * Nvl(rsTemp!�����ܷ���, 0) & "," & -1 * Nvl(rsTemp!ҽ���ܷ���, 0) & ")"
        gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    סԺ�������_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Sub ��ѯǷ�ѵ�λ_����(ByVal str��λ���� As String, ByVal str������� As String)
'���ܣ����ýӿڲ�ѯǷ�ѵ�λ
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim str��ʾ As String
    
    If str��λ���� = "" Then Exit Sub
'    str��λ���� = String(12 - Len(str��λ����), "0") & str��λ����
    
    On Error GoTo errHandle
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "DEPTCODE", str��λ����)                '��λ����
    Call InsertChild(mdomInput.documentElement, "INSURETYPE", str�������)              '�������
    
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
            Case "10"
                str��ʾ = str��ʾ & "������Ա����"
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

Public Function ҽ����Ŀ_����(rsTemp As ADODB.Recordset, Optional ByVal str��� As String = "12") As Boolean
'���ܣ�ҽ������ҩƷĿ¼��ѯ
'��ǰ�����Ĳ�ѯ���ָ�Ϊ����Ŀ֧������ѯ��41-�������� 42-����סԺ 21-�������� 22-����סԺ 11-��ͨ���� 12-��ͨסԺ 31-�������� 32-����סԺ��
'����ͨסԺ������Ŀ�嵥�����գ�����ģʽ����ǰһ����ֻ���ṩ����ѯ�Ľ��棬�ɰ��û�Ҫ���ѯĳ������µ���Ŀ��֧����������Ϣ
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim str���� As String, str���� As String, str����, str��ע As String
    Dim str��ʼ���� As String, str�������� As String, str��ǰ���� As String
        
    On Error GoTo errHandle
    
    str��ǰ���� = Format(zldatabase.Currentdate(), "yyyy-MM-dd")
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "ITEMCODE", "")         ' ҽ������
    Call InsertChild(mdomInput.documentElement, "ITEMPAYTYPE", str���) ' ��Ŀ֧�����
    
    '���ýӿ�
    If CommServer("QUERYSERVICE") = False Then Exit Function
    
    Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
    If nodRowset Is Nothing Then Exit Function
    For Each nodRow In nodRowset.childNodes
        str���� = GetAttributeValue(nodRow, "ITEMCODE")
        str���� = ToVarchar(Replace(GetAttributeValue(nodRow, "ITEMNAME"), "'", ""), 40)
        str���� = ToVarchar(zlcommfun.SpellCode(str����), 10)
        str��ʼ���� = Mid(GetAttributeValue(nodRow, "STARTDATE"), 1, 10)
        str�������� = Mid(GetAttributeValue(nodRow, "ENDDATE"), 1, 10)
'        PRICELMT           '����޼�
'        SELFRATE           '�Ը�����
'        BEARINGITEMFLAG    '������Ŀ��־
'        GSITEMFLAG         '������Ŀ��־
'        SPECPAYFLAG        '���ⱨ����Ŀ��־
'        BGITEMTYPE         '���ɽ�����Ŀ���
        str��ע = Val(GetAttributeValue(nodRow, "PRICELMT")) & "|" & Val(GetAttributeValue(nodRow, "SELFRATE")) & "|" & _
                  Val(GetAttributeValue(nodRow, "BEARINGITEMFLAG")) & "|" & Val(GetAttributeValue(nodRow, "GSITEMFLAG")) & "|" & _
                  Val(GetAttributeValue(nodRow, "SPECPAYFLAG")) & "|" & Val(GetAttributeValue(nodRow, "BGITEMTYPE"))
        
        If str���� <> "" And str��ǰ���� >= str��ʼ���� And str��ǰ���� <= str�������� Then
            rsTemp.AddNew Array("CLASSCODE", "CODE", "NAME", "PY", "MEMO"), Array("1", str����, str����, str����, str��ע)
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

Public Function CommRecServer(ByVal strFunction As String) As Boolean
'���ܣ�����ҽ������
    Dim InvokeServer As String '����ǰ�÷������ķ���ֵ
    Dim StrInput As String
    
    '�����Ĵ���
    StrInput = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?>" & vbCrLf & mdomInput.xml
    Call DebugTool(StrInput)
    
    Select Case strFunction
        Case "APPRECM"
            InvokeServer = obj����.APPRECM("ZFRJ", StrInput)
        Case "DELRECM"
            InvokeServer = obj����.DELRECM("ZFRJ", StrInput)
        Case "APPRECB"
            InvokeServer = obj����.APPRECB("ZFRJ", StrInput)
        Case "DELRECB"
            InvokeServer = obj����.DELRECB("ZFRJ", StrInput)
        Case "QUERYREC"
            InvokeServer = obj����.QUERYREC("ZFRJ", StrInput)
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
            CommRecServer = True
        Else
            '����ʧ��
            InvokeServer = GetElemnetValue("INFO")
            If InvokeServer = "" Then InvokeServer = "����������ʧ�ܡ�"
            MsgBox "ҽ�����������ش���" & vbCrLf & vbCrLf & InvokeServer, vbInformation, gstrSysName
        End If
    End If
End Function

Public Function CommServer(ByVal strFunction As String, Optional ByVal strAdvance As String = "") As Boolean
'���ܣ�����ҽ������
    Dim InvokeServer As String '����ǰ�÷������ķ���ֵ
    Dim StrInput As String
    
    '�����Ĵ���
    StrInput = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?>" & vbCrLf & mdomInput.xml
    Call DebugTool(StrInput)
    
    Select Case strFunction
        Case "GETPSNINFO"
            InvokeServer = objҽ��.GETPSNINFO("ZFRJ", StrInput)
        Case "MODIFYCARD"               '�޸Ŀ�����
            InvokeServer = objҽ��.MODIFYCARD("ZFRJ", StrInput)
        Case "GETCLINNO"                '����Һ�
            InvokeServer = objҽ��.GETCLINNO("ZFRJ", StrInput)
        Case "CALCLIN"                  '��ͨ����֧��
            InvokeServer = objҽ��.CALCLIN("ZFRJ", StrInput)
        Case "CALSPECCLIN"              '��������֧��
            InvokeServer = objҽ��.CALSPECCLIN("ZFRJ", StrInput)
        Case "RETBALANCE"               '��Ʊ
            If strAdvance = "1" Then    '������Ʊ
                InvokeServer = objҽ��.RETLX("ZFRJ", StrInput)
            Else
                InvokeServer = objҽ��.RETBALANCE("ZFRJ", StrInput)
            End If
        Case "HOSPREG"                  'סԺ�Ǽ�
            InvokeServer = objҽ��.HOSPREG("ZFRJ", StrInput)
        Case "HOSPOUT"                  '��Ժ�Ǽ�
            InvokeServer = objҽ��.HOSPOUT("ZFRJ", StrInput)
        Case "CALHOSP"                  'סԺ֧��
            If strAdvance = "1" Then    '�޿����㣬�������ڲ����ӷѵ����
                InvokeServer = objҽ��.CALHOSPSP("ZFRJ", StrInput)
            Else
                InvokeServer = objҽ��.CALHOSP("ZFRJ", StrInput)
            End If
        Case "SETRECKONINGTYPE"         '�������㷽ʽ
            InvokeServer = objҽ��.SETRECKONINGTYPE("ZFRJ", StrInput)
        Case "QUERYHOSPSINGLEILLNESS"   '��������������
            InvokeServer = objҽ��.QUERYHOSPSINGLEILLNESS("ZFRJ", StrInput)
        Case "QUERYHOSPSINGLEILLNESS_BG"   '�����ֽ���Ŀ¼
            InvokeServer = objҽ��.QUERYHOSPSINGLEILLNESS_BG("ZFRJ", StrInput)
        Case "QUERYSERVICE"              'ҽ������ҩƷĿ¼��ѯ
            InvokeServer = objҽ��.QUERYSERVICE("ZFRJ", StrInput)
        Case "QUERYARREARDEPT"          '��ѯǷ�ѵ�λ
            InvokeServer = objҽ��.QUERYARREARDEPT("ZFRJ", StrInput)
        Case "GETHOSPSINGLEILLNESS"     '���ص�������������
            InvokeServer = objҽ��.GETHOSPSINGLEILLNESS("ZFRJ", StrInput)
        Case "GETHOSPSINGLEILLNESS_BG"  '���ص����ֽ���Ŀ¼
            InvokeServer = objҽ��.GETHOSPSINGLEILLNESS_BG("ZFRJ", StrInput)
        Case "SETBEARINGFLAG"           '����������־
            InvokeServer = objҽ��.SETBEARINGFLAG("ZFRJ", StrInput)
        Case "UPLOADICD"                '�ϴ�ICD����
            InvokeServer = objҽ��.UPLOADICD("ZFRJ", StrInput)
        Case "SETCALTYPE"
            InvokeServer = objҽ��.SETCALTYPE("ZFRJ", StrInput)
        Case "RETHOSPOUT"
            InvokeServer = objҽ��.RETHOSPOUT("ZFRJ", StrInput)
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

Public Function Get��֤_����(bytType As Byte, str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, _
                ByVal lng����ID As Long, Optional blnǿ��ˢ�� As Boolean = False) As Boolean
'���ܣ��õ�ҽ�����˵Ļ������������֤��Ϣ
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    If blnǿ��ˢ�� = False And lng����ID > 0 Then
        '�����ݿ��ж����Ѵ洢��ֵ
        gstrSQL = "select ����,ҽ����,˳���,���� from �����ʻ� where ����ID=" & lng����ID & " and ����=" & TYPE_������
        Call OpenRecordset(rsTemp, "����ҽ��")
        
        If rsTemp.EOF = False Then
            strTemp = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            strTemp = Replace(strTemp, ",", ";")
            If strTemp = mstr���� And mstr���� <> "" Then
                '��ͬһ����
                str���� = mstr����
                str���� = mstr����
            Else
                str���� = strTemp
                str���� = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            End If
            
            strҽ���� = IIf(IsNull(rsTemp("ҽ����")), "", rsTemp("ҽ����"))
            str�����ı�� = mstrҽ�����ı���_����
            
            Get��֤_���� = True
            Exit Function
        End If
    End If
    
    If frmIdentify����.GetIdentify(bytType, str����, strҽ����, str�����ı��, str����) = False Then
        Exit Function
    Else
        'ˢ����Ȼ��ȷ����Ҫ����Ƿ���ǵ�ǰ���˵�
            str���� = Split(str����, "^")(0)
            If lng����ID > 0 Then
            '�����ݿ��ж����Ѵ洢��ֵ
            gstrSQL = "select ����,ҽ����,˳��� from �����ʻ� where ����ID=" & lng����ID & " and ����=" & TYPE_������
            Call OpenRecordset(rsTemp, "����ҽ��")
            
            If str���� <> Replace(IIf(IsNull(rsTemp("����")), "", rsTemp("����")), ",", ";") Or strҽ���� <> IIf(IsNull(rsTemp("ҽ����")), "", rsTemp("ҽ����")) Then
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
             " Where A.���>=6 And A.����=" & TYPE_������ & " And A.������=B.���� And B.����='" & strData & "'"
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
    Dim curҽ������ As Currency, cur���ͳ�� As Currency, cur����Ա���� As Currency
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, str����˳��� As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    gstrSQL = "Select B.����ID,B.����,B.ҽ����,B.���� " & _
        " From ����Ԥ����¼ A,�����ʻ� B " & _
        " Where A.����ID=B.����ID And B.����=" & TYPE_������ & _
        "       And A.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "ȡ����ID")
    If rsTemp.EOF Then Exit Function
    lng����ID = rsTemp!����ID
    If Get��֤_����(0, str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
    
    datCurr = zldatabase.Currentdate()
    
    'ȡ�ʻ����
    gstrSQL = "Select Nvl(�ʻ����,0) �ʻ���� From �����ʻ� Where ����=" & TYPE_������ & " And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "ȡ�ʻ����")
    cur�ʻ���� = rsTemp!�ʻ����
    
    '��XML����ֵ
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", strҽ����)     ' ���˱���
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) ' ��������
    
    '���ýӿ�
    If CommServer("GETCLINNO") = False Then Exit Function
    str����˳��� = GetElemnetValue("BILLNO")
    
    gstrSQL = "Select ����ID,�շ�ϸĿID,����*NVL(����,1) AS ����,��׼���� AS ����,Nvl(ʵ�ս��,0) AS ʵ�ս��,���ձ���,'  ' AS ժҪ" & _
        " From ���˷��ü�¼ " & _
        " Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Call OpenRecordset(rsTemp, "����ҽ��")
    If Not �����������_����(rsTemp, str���㷽ʽ, "") Then Exit Function
    
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
        Case "ҽ�Ʋ���"
            cur����Ա���� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        End Select
    Next
    
    If Not �������_����(lng����ID, cur�����ʻ�, "") Then Exit Function
    
   '��Ҫ����������
    str���㷽ʽ = ""
    If cur�����ʻ� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & cur�����ʻ�
    If curҽ������ <> 0 Then str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & curҽ������
    If cur���ͳ�� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||���ͳ��|" & cur���ͳ��
    If cur����Ա���� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||ҽ�Ʋ���|" & cur����Ա����
    If str���㷽ʽ <> "" Then
        str���㷽ʽ = Mid(str���㷽ʽ, 3)
        gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',0)"
        Call zldatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
    End If
    
    ����Һ�_���� = True
    
    Call frm������Ϣ.ShowME(lng����ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ���ý��㷽ʽ_����(ByVal lng����ID As Long, ByVal frmParent As Object, Optional ByVal blnסԺ As Boolean = False) As String
    '���ؽ��㷽ʽ�뵥���ֱ���
    ���ý��㷽ʽ_���� = frm���ý��㷽ʽ.ShowSelect(lng����ID, TYPE_������, blnסԺ, frmParent)
End Function

Public Function �������㷽ʽ_����(ByVal lng����ID As Long, ByVal frmParent As Object) As Boolean
    �������㷽ʽ_���� = frm�������㷽ʽ.ShowSelect(lng����ID, TYPE_������, frmParent)
End Function

Public Sub ����ѡ��_����(ByVal lng����ID As Long)
    Dim lng����ID As Long
    Dim str���� As String
    Dim rs���� As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    '��ȡ�ò�����ǰ�Ĳ�����Ϣ
    gstrSQL = " select B.����,B.���� from �����ʻ� A,���ղ��� B " & _
              " where A.����ID=" & lng����ID & " and A.����=" & TYPE_������ & " and A.����ID=B.ID"
    Call OpenRecordset(rsTemp, "סԺԤ��")
    If rsTemp.RecordCount <> 0 Then
        str���� = "[" & rsTemp!���� & "]" & rsTemp!����
    End If
    
    '��סԺ����ѡ����
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & TYPE_������
    Set rs���� = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ�����֡���" & IIf(str���� = "", "��", str����))
    If Not rs���� Is Nothing Then
        lng����ID = rs����("ID")
    End If
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'����ID','''" & lng����ID & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "���没��")
End Sub

Public Function ����ICD����_����(ByVal lng����ID As Long) As Boolean
    Dim strICD As String
    Dim rsTemp As New ADODB.Recordset
'    <BILLNO>����˳���</BILLNO>
'    <ICD>ICD����</ICD>
'    <DODATE>����ʱ��</DODATE>
    
    If Not ҽ�������Ѿ���Ժ(lng����ID) Then
        MsgBox "��ҽ�����˻�δ��Ժ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ѡ��ICD����
    strICD = frm����ѡ��_����.ChooseDisease(lng����ID)
    If strICD = "" Then Exit Function
    
    '�ϴ����˵�ICD����
    gstrSQL = "Select ҽ����,˳��� From �����ʻ� Where ����=" & TYPE_������ & " ANd ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "ȡ�ò��˵�ҽ����")
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "BILLNO", Nvl(rsTemp!˳���))   '˳���
    Call InsertChild(mdomInput.documentElement, "ICD", strICD)                  '����
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) '��������
    If CommServer("UPLOADICD") = False Then Exit Function
    
    ����ICD����_���� = True
End Function

Public Function GetItemInsure_����(lng����ID As Long, lng�շ�ϸĿID As Long, bln���� As Boolean) As String
    'ҽ����������в���һ����¼
    'insert into ҽ���������
    '(����,����,����,˵��)
    'Values
    '(50,'1','����','��')
    '����ʷ�������ݲ��뵽ҽ��������ϸ��
    'insert into ҽ��������ϸ
    'select ����,1,�շ�ϸĿID,��Ŀ����,''
    'From ����֧����Ŀ
    'Where ���� = 50
    Dim strDefault As String            'ȱʡҽ������
    Dim strCurrent As String            '��ǰҽ�����룬����ȡ������룬סԺȡסԺ����
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = "Select B.���,A.����,A.����,B.˵�� From ������Ŀ A,ҽ��������ϸ B" & _
        " Where B.����=" & TYPE_������ & " And A.����=B.���� And A.����=B.��Ŀ���� And B.�շ�ϸĿID=" & lng�շ�ϸĿID
    Call OpenRecordset(rsTemp, "��ȡҽ������")
    rsTemp.Filter = "���=1"
    Select Case rsTemp.RecordCount
    Case 0
        'û�����ö�Ӧ���룬ȡȱʡ����
        rsTemp.Filter = "���=0"
        If rsTemp.RecordCount <> 0 Then
            GetItemInsure_���� = rsTemp!����
        End If
    Case 1
        GetItemInsure_���� = rsTemp!����
    Case Else
        '��ѡ
        GetItemInsure_���� = frmҽ����Ŀѡ��.ShowSelect(rsTemp, lng�շ�ϸĿID)
    End Select
    
    rsTemp.Filter = 0
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    rsTemp.Filter = 0
End Function

Private Function IS����(ByVal lng����ID As Long) As Boolean
    Dim str��Ա��� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = " Select ��Ա��� From �����ʻ� Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ��Ա���")
    If rsTemp.RecordCount = 0 Then Exit Function
    str��Ա��� = Nvl(rsTemp!��Ա���)
    str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21" _
                  , str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", True, "11")
    IS���� = (str��Ա��� = "32" Or str��Ա��� = "34")
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetNextID(ByVal strTable As String, ByVal cnCustom As ADODB.Connection) As Long
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select " & strTable & "_ID.Nextval From Dual"
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, cnCustom
    GetNextID = rsTemp.Fields(0).Value
End Function
