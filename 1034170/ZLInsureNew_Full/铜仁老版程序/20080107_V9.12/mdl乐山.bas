Attribute VB_Name = "mdl��ɽ"
Option Explicit
Public Declare Sub LS_ErrMessage Lib "SIHisInterface.dll" Alias "GetErrorMessage" (ErrorMsg As TStringOfChar)
Public Declare Function LS_UserLogin Lib "SIHisInterface.dll" Alias "UserLogin" (UserCode As TStringOfChar, PWD As TStringOfChar) As Byte
Public Declare Function LS_ChangePwd Lib "SIHisInterface.dll" Alias "ChangeUserPwd" (OldPwd As TStringOfChar, NewPWD As TStringOfChar) As Byte
Public Declare Sub LS_UserLogout Lib "SIHisInterface.dll" Alias "UserLogout" ()
Public Declare Function LS_ConnectServer Lib "SIHisInterface.dll" Alias "ConnectServer" (ServerName As TStringOfChar) As Byte
Public Declare Sub LS_DisConnectServer Lib "SIHisInterface.dll" Alias "DisConnectServer" ()

'��ȡ�α�����Ϣ
Public Declare Function LS_GetPersonInfo Lib "SIHisInterface.dll" Alias "GetPersonInfo" (PInfo As �����Ϣ) As Byte
'��Ժ�Ǽ�
Public Declare Function LS_InHospitalRegister Lib "SIHisInterface.dll" Alias "InBedRegster" (InBedRegInfo As סԺ��Ϣ) As Byte
'��ȡ��Ժ�Ǽ���Ϣ
Public Declare Function LS_GetInHospitalRegInfo Lib "SIHisInterface.dll" Alias "GetInBedRegInfo" (InBedRegID As TStringOfChar) As Byte
'¼��ҩƷ����
Public Declare Function LS_AddDrug Lib "SIHisInterface.dll" Alias "AddDrug" (DrugInfo As ҩƷ��Ϣ) As Byte
'¼�����Ʒ���
Public Declare Function LS_AddDiag Lib "SIHisInterface.dll" Alias "AddDiag" (DiagInfo As ������Ϣ) As Byte
'¼�������ʩ����
Public Declare Function LS_AddService Lib "SIHisInterface.dll" Alias "AddServiceItem" (ServiceItemInfo As ������ʩ��Ϣ) As Byte
'���������ϸ
Public Declare Function LS_SaveDetail Lib "SIHisInterface.dll" Alias "InBedRegApplyUpdates" (InBedRegID As TStringOfChar) As Byte
'סԺ����Ԥ����
Public Declare Function LS_PreBalance Lib "SIHisInterface.dll" Alias "NewInBedBill" (InBedBillInfo As סԺ������Ϣ) As Byte
'סԺ���ý���
Public Declare Function LS_Balance Lib "SIHisInterface.dll" Alias "SaveInBedBill" (InBedBillInfo As סԺ������Ϣ) As Byte

'ȫ�ֱ�����
Private Const mstr��Ժ���� As String = "��Ժ����"
Private Const mstr��;�ݽ��� As String = "��;�ݽ���"
Private Const mstrתԺ���� As String = "תԺ����"

'���������Ϣ����
Private Const ��Ժ���ұ�� = 0
Private Const ��Ժ�������� = 1
Private Const ��Ժ������� = 2
Private Const ��Ժ�������� = 3
Private Const ��Ժ������� = 4
Private Const ��Ժ�������� = 5
Private Const סԺҽʦ = 6
Private Const סԺ�� = 7
Private Const ��Ժ��� = 8
Private Const ��Ժ��� = 9
Private Const ��Ժ���� = 10
Private Const ��Ժ��ʽ = 11

Public Type TStringOfChar
    Data As String * 100
End Type
Public Type �����Ϣ                   'TPersonInfo
    '��������Ϊ��������
    PSN_ID              As Long      'ҽ�Ʋα�ID��
    PSN_No              As Long      '�α��˱���
    PSN_NAME            As String * 100 '�α�������
    Sex                 As String * 100 '�Ա�
    IDCARD              As String * 100 '���֤����
    PSN_STS             As String * 100 '�α���״̬
    PSN_TYP             As String * 100 '��Ա���
    UNIT_CODE           As String * 100 '��λ����
    UNIT_NAME           As String * 100 '��λ����
    OFFICAL_TYP         As String * 100 '����Ա���
    HAI_TYP             As String * 100 '����ҽ������
    ACCT_STS            As String * 100 'ҽ���˻�״̬
    HI_ACCT_PWD         As String * 100 'ҽ���ʻ�����
    SILL_PAY_AMT_TOTAL  As Single       '���ڽ����������⼲��֧�����
    SILL_YR_FUND_AMT    As Single       '��������ͳ�����֧�����
    YR_FUND_AMT         As Single       '����ͳ�����֧�����
    HAI_YR_HIGH_AMT     As Single       '���ڲ���߶�֧�����
    HAI_YR_INBED_AMT    As Single       '���ڲ���סԺ����֧�����
    GZ_CUR_AMT          As Single       '�����˻����
    YR_INBED_CNT        As Long      '����סԺ����
End Type
Private Type סԺ��Ϣ                   'TInBedRegInfo
    PSN_ID              As Long      'ҽ�Ʋα���ID��
    INBED_SILL_ID       As Long      'סԺ���ⲡ��ID��������
    INBED_NO            As String * 100 'סԺ��
    INBED_EXAM          As String * 100 '��Ժ���
    INBED_EXAM_ICD10_NO As String * 100 '��Ժ���ICD10����
    INBED_DEPT          As String * 100 '��Ժ����
    '��������Ϊ��������
    INBED_REG_ID        As String * 100 'סԺ�Ǽ�ID
    INBED_DT            As String * 100 '��Ժʱ�䣬¼������
End Type
Private Type ҩƷ��Ϣ               'TDrugInfo
    INBED_REG_ID    As String * 100 'סԺ�Ǽ�ID
    RECEIPT_DT      As String * 100 '�շ�ʱ��
    DRUG_CATALOG_ID As String * 100 'ҩƷ�������ID
    DRUG_INFO       As String * 100 'ҩƷ��Ϣ
    UNIT_PRC        As Single       '����
    SRVC_CNT        As Single       '����
    COST_PRC        As Single       '�ɱ�����
    DRUG_TYP        As String * 100 'ҩ�����
    DRUG_SPEC       As String * 100 'ҩ����
    PRODUCE_FACTORY As String * 100 '��������
    '��������Ϊ��������
    FEE_ITEM_TYP    As String * 100 '������Ŀ����
    FEE_TYP         As String * 100 '��������
    PART_PUB_AMT    As Single       '���ֹ��ѽ��
    PART_SELF_AMT   As Single       '�����Էѽ��
    PUB_PAY_AMT     As Single       '���ѽ��
    SELF_PAY_AMT    As Single       '�Էѽ��
    SELF_PAY_PCT    As Single       '�Էѱ���
    MAX_RETAIL_PRC  As Single       '������ۼ�
End Type
Private Type ������Ϣ               'TDiagInfo
    INBED_REG_ID    As String * 100 'סԺ�Ǽ�ID
    RECEIPT_DT      As String * 100 '�շ�ʱ��
    DIAG_CATALOG_ID As String * 100 '������Ŀ�������ID
    DIAG_ITEM_NAME  As String * 100 '������Ŀ����
    UNIT_PRC        As Single       '����
    SRVC_CNT        As Single       '����
    '��������Ϊ��������
    FEE_ITEM_TYP    As String * 100 '������Ŀ����
    FEE_TYP         As String * 100 '��������
    PART_PUB_AMT    As Single       '���ֹ��ѽ��
    PART_SELF_AMT   As Single       '�����Էѽ��
    PUB_PAY_AMT     As Single       '���ѽ��
    SELF_PAY_AMT    As Single       '�Էѽ��
    SELF_PAY_PCT    As Single       '�Էѱ���
    MAX_RETAIL_PRC  As Single       '������ۼ�
End Type
Private Type ������ʩ��Ϣ           'TServiceItemInfo
    INBED_REG_ID    As String * 100 'סԺ�Ǽ�ID
    RECEIPT_DT      As String * 100 '�շ�ʱ��
    SRVC_ITEM_ID    As String * 100 '����ҽ�Ʊ��շ�����ʩ��׼
    SRVC_NAME       As String * 100 '������ʩ����
    UNIT_PRC        As Single       '����
    SRVC_CNT        As Single       '����
    '��������Ϊ��������
    FEE_ITEM_TYP    As String * 100 '������Ŀ����
    FEE_TYP         As String * 100 '��������
    PART_PUB_AMT    As Single       '���ֹ��ѽ��
    PART_SELF_AMT   As Single       '�����Էѽ��
    PUB_PAY_AMT     As Single       '���ѽ��
    SELF_PAY_AMT    As Single       '�Էѽ��
    SELF_PAY_PCT    As Single       '�Էѱ���
    MAX_RETAIL_PRC  As Single       '������ۼ�
End Type
Private Type סԺ������Ϣ                   'TInBedBillInfo
    INBED_REG_ID        As String * 100     'סԺ�Ǽ�ID
    EXAM_TYP            As String * 100     '�������
    INBED_STL_TYP       As String * 100     'סԺ���ʷ�ʽ
    OUTBED_EXAM         As String * 100     '��Ժ���
    OUTBED_EXAM_ICD10_NO As String * 100    '��Ժ���ICD10����
    OUTBED_DEPT         As String * 100     '��Ժ����
    ILL_TRS_STS         As String * 100     '����ת��(������������)
    INBED_DOCTOR        As String * 100     '�ܴ�ҽ��
    OUTBED_DT           As String * 100     '��Ժʱ��
    '��������Ϊ��������
    INBED_DAY_CNT       As Long          'סԺ����
    FEE_STL_LOC         As String * 100     '���ý���ص�
    EXAM_ADDR           As String * 100     '����ص�
    INBED_STL_BILL_ID   As String * 100     'סԺ���ʵ�id
    INBED_STL_BILL_NO   As String * 100     'סԺ���ʵ���
    PART_PUB_AMT        As Single           '���ֹ��ѽ��
    PART_SELF_AMT       As Single           '�����Էѽ��
    PUB_PAY_AMT         As Single           '���ѽ��
    SELF_PAY_AMT        As Single           '�Էѽ��
    INBED_FUND_AMT      As Single           'סԺͳ��֧�����
    INBED_ACCT_AMT      As Single           'סԺ����֧�����
    CASH_PAY_AMT        As Single           '�ֽ�֧�����
    HAI_INBED_SBS_AMT   As Single           '����סԺ����֧�����
    HAI_INBED_AMT       As Single           '����סԺ֧�����
    HAI_INBED_REPAY_AMT As Single           '����סԺ�ٴ�֧�����
    HAI_INBED_HIGH_AMT  As Single           '����סԺ�߶�֧�����
    OFFICAL_HIGH_AMT    As Single           '����Ա�߶��֧�����
    OFFICAL_INBED_AMT   As Single           '����ԱסԺ����֧�����
    OFFICAL_ACCT_AMT    As Single           '����Ա���ʲ���֧�����
End Type
Private Type ������Ϣ
    ˳��� As TStringOfChar
    �ܷ��� As Currency
    �ֽ� As Currency
    �����ʻ� As Currency
    ҽ������ As Currency
    ������� As Currency
End Type
Public gPersonInfo_��ɽ As �����Ϣ
Public gInBedRegInfo_��ɽ As סԺ��Ϣ
Public gDrugInfo_��ɽ As ҩƷ��Ϣ
Public gDiagInfo_��ɽ As ������Ϣ
Public gServiceItemInfo_��ɽ As ������ʩ��Ϣ
Public gInBedBillInfo_��ɽ As סԺ������Ϣ
Private gtypBalance As ������Ϣ

Private glngInterface_��ɽ As Long
Private gstrErrMsg_��ɽ As TStringOfChar          '������Ϣ
Public gbytReturn_��ɽ As Byte                '0-����;����ֵ��������

Public Function ҽ����ʼ��_��ɽ() As Boolean
    Dim strServer As TStringOfChar
    On Error GoTo ErrHand
    
    If glngInterface_��ɽ <> 0 Then ҽ����ʼ��_��ɽ = True: Exit Function
    strServer = GetServerInfo
    If strServer.Data = "" Then Exit Function
    
    '���ӷ�����
    gbytReturn_��ɽ = LS_ConnectServer(strServer)
    If GetErrInfo_��ɽ Then Exit Function
    
    '��¼����(ʧ����Ͽ����Ӳ��˳�)
    If Not frm��¼����.LoginCenter(TYPE_��ɽ, True) Then
        Call ҽ����ֹ_��ɽ
        Exit Function
    End If
    glngInterface_��ɽ = 1
    
    ҽ����ʼ��_��ɽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
  Resume
    End If
End Function

Public Function ҽ����ֹ_��ɽ() As Boolean
    On Error Resume Next
    If glngInterface_��ɽ = 0 Then
        ҽ����ֹ_��ɽ = True
        Exit Function
    End If
    
    '����Ա�˳�
    Call LS_UserLogout
    '���ӷ�����
    Call LS_DisConnectServer
    glngInterface_��ɽ = 0
    
    ҽ����ֹ_��ɽ = True
End Function

Public Function ҽ������_��ɽ() As Boolean
    With frmSet��ɽ
        ҽ������_��ɽ = .ShowME
    End With
End Function

Public Function GetErrInfo_��ɽ() As Boolean
    If gbytReturn_��ɽ = 1 Then Exit Function
    Call LS_ErrMessage(gstrErrMsg_��ɽ)
    MsgBox gstrErrMsg_��ɽ.Data, vbInformation, gstrSysName
    GetErrInfo_��ɽ = True
End Function

Private Function GetServerInfo() As TStringOfChar
    '��ȡ��������ַ
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    '��ȡ��������ַ���˿ڼ��������('��������ַ','�������˿ں�','��������ڳ���')
    gstrSQL = " Select ������,����ֵ From ���ղ���" & _
              " Where ����=" & TYPE_��ɽ & " And ������ = '��������ַ'"
    Call OpenRecordset(rsTemp, "��ȡ���������ƻ�IP��ַ")
    
    With rsTemp
        If .RecordCount = 0 Then Exit Function
        GetServerInfo.Data = NVL(!����ֵ)
    End With
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �������_��ɽ(ByVal strҽ���� As String) As Currency
    '����: ֱ�Ӷ������ڽ��
    '����: �Ƿ����
    '����: ���ظ����ʻ����
    Dim rsAccount As New ADODB.Recordset
    On Error GoTo ErrHand
    
    gstrSQL = " Select Nvl(�ʻ����,0) �ʻ���� From �����ʻ� " & _
              " Where ����=" & gintInsure & " And ҽ����='" & strҽ���� & "'"
    Call OpenRecordset(rsAccount, "���ظ����ʻ����")
    
    �������_��ɽ = rsAccount!�ʻ����
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �����������_��ɽ(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    'cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    On Error GoTo ErrHand
    
    If str���㷽ʽ = "" Then str���㷽ʽ = "�ֽ�;0;0"
    �����������_��ɽ = False
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �������_��ɽ(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    'cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    'ע�⣺�ӿڹ涨��������ϸ�������ϴ���סԺ��ϸ��Ԥ����ʱ�ϴ���������ڽ��㣬����ʹ��Ȧ��ӿڣ����������Ǯ���������ڣ������ӿ��ڽ��
    '���������Ҫͨ��������������ȡ����Ȧ�����ǽӿڷ��أ���Ҫ�޸�
    On Error GoTo ErrHand
    
    �������_��ɽ = False
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ����������_��ɽ(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    'cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    On Error GoTo ErrHand
    
    ����������_��ɽ = False
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_��ɽ(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    Dim str˳��� As String
    Dim arrPatient
    On Error GoTo ErrHand
    
    arrPatient = Split(��ȡ���������Ϣ(lng����ID, lng��ҳID), "||")
    'д�������
    With gInBedRegInfo_��ɽ
        .PSN_ID = gPersonInfo_��ɽ.PSN_ID                           'סԺ�α�ID��
        .INBED_SILL_ID = 0                                          'סԺ���ⲡ��ID��������
        .INBED_NO = arrPatient(סԺ��)                              'סԺ��
        .INBED_EXAM = Split(arrPatient(��Ժ���), "|")(0)           '��Ժ���
        .INBED_EXAM_ICD10_NO = Split(arrPatient(��Ժ���), "|")(1)  '��Ժ���ICD10����
        .INBED_DEPT = arrPatient(��Ժ��������)                          '��Ժ����
    End With
    
    '������Ժ�Ǽǽӿ�
    gbytReturn_��ɽ = LS_InHospitalRegister(gInBedRegInfo_��ɽ)
    If GetErrInfo_��ɽ Then Exit Function
    
    '���¸����ʻ��е���Ϣ
    str˳��� = TrimTsChar(gInBedRegInfo_��ɽ.INBED_REG_ID)
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & gintInsure & ",'˳���','''" & str˳��� & "''')"
    Call ExecuteProcedure("������Ժҵ�����к�")
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure("������Ժ�Ǽ�")

    ��Ժ�Ǽ�_��ɽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_��ɽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ��
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    '����������Ժ
    On Error GoTo ErrHand
    
    MsgBox "��֧�ֳ�Ժ�Ǽǳ���������ҽ���ӿ�����ϵ��", vbInformation, gstrSysName
    ��Ժ�Ǽǳ���_��ɽ = False
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_��ɽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    On Error GoTo ErrHand
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ��
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false

    '����HIS��Ժ
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure("��Ժ�Ǽ�")

    ��Ժ�Ǽ�_��ɽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽǳ���_��ɽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    On Error GoTo ErrHand

    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure("��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_��ɽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function סԺ�������_��ɽ(rsExse As Recordset, ByVal lng����ID As Long) As String
    Dim lng��ҳID As Long
    Dim bln��Ժ���� As Boolean
    Dim str��¼���� As String, str��¼״̬ As String, strNO As String
    Dim arrPatient
    Dim rsTemp As New ADODB.Recordset
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    '�ӿڷ��صı������ȥ����סԺ�ڼ�����������Ļ��ܽ��󣬲��Ǳ��ε�ʵ�ʱ�����
    'rsExse��¼���е��ֶ��嵥
    '��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID,
    '�շ����,�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������
    '���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��,�Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    On Error GoTo ErrHand
    
    '��ȡ�ܷ���
    gtypBalance.�ܷ��� = 0
    With rsExse
        Do While Not .EOF
            gtypBalance.�ܷ��� = gtypBalance.�ܷ��� + NVL(!���, 0)
            '�ϴ���ϸ
            If NVL(!�Ƿ��ϴ�, 0) = 0 And (strNO <> !NO Or str��¼���� <> !��¼���� Or str��¼״̬ <> !��¼״̬) Then
                strNO = !NO
                str��¼���� = !��¼����
                str��¼״̬ = !��¼״̬
                If Not �ϴ�����_��ɽ(str��¼����, str��¼״̬, strNO) Then Exit Function
            End If
            .MoveNext
        Loop
    End With
    
    '��ȡ��ҳID
    gstrSQL = "Select סԺ���� ��ҳID From ������Ϣ Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ��ҳID")
    lng��ҳID = rsTemp!��ҳID
    
    Call ��ȡ���˻�����Ϣ(lng����ID)
    arrPatient = Split(��ȡ���������Ϣ(lng����ID, lng��ҳID), "||")
    bln��Ժ���� = ҽ�������Ѿ���Ժ(lng����ID)
    
    'д�������
    With gInBedBillInfo_��ɽ
        .INBED_REG_ID = gtypBalance.˳���.Data
        .EXAM_TYP = ""
        .INBED_STL_TYP = IIf(bln��Ժ����, IIf(arrPatient(��Ժ��ʽ) = "תԺ", mstrתԺ����, mstr��Ժ����), mstr��;�ݽ���)
        .OUTBED_EXAM = Split(arrPatient(��Ժ���), "|")(0)
        .OUTBED_EXAM_ICD10_NO = Split(arrPatient(��Ժ���), "|")(1)
        .OUTBED_DEPT = arrPatient(��Ժ��������)
        .ILL_TRS_STS = "����"
        .INBED_DOCTOR = arrPatient(סԺҽʦ)
        .OUTBED_DT = IIf(bln��Ժ����, arrPatient(��Ժ����), "")
    End With
    gbytReturn_��ɽ = LS_PreBalance(gInBedBillInfo_��ɽ)
    If GetErrInfo_��ɽ Then Exit Function

    Call Get������Ϣ
    סԺ�������_��ɽ = "�����ʻ�;" & gtypBalance.�����ʻ� & ";0"
    סԺ�������_��ɽ = סԺ�������_��ɽ & "|ҽ������;" & gtypBalance.ҽ������ & ";0"
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_��ɽ(lng����ID As Long, ByVal lng����ID As Long) As Boolean
    Dim cur�ʻ�֧�� As Currency
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
  '������㣨���ص����ݼ�ȥ���ν������ݣ��͵��ڱ��ε���ʵ�������ݣ�
    On Error GoTo ErrHand
    Call ��ȡ���˻�����Ϣ(lng����ID)
    
    '��ȡ���θ����ʻ�֧����
    gstrSQL = "Select Nvl(A.��Ԥ��,0) �����ʻ� " & _
        " From ����Ԥ����¼ A,�����ʻ� B " & _
        " Where A.����ID=B.����ID And B.����=" & gintInsure & _
        " And A.���㷽ʽ in ('�����ʻ�') And A.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ���θ����ʻ�֧����")
    cur�ʻ�֧�� = 0
    If Not rsTemp.EOF Then
        cur�ʻ�֧�� = rsTemp!�����ʻ�
    End If
    
    'ֱ�ӵ��ý���ӿڣ���Ϊ��������Ѿ���д����ڲ���
    gbytReturn_��ɽ = LS_Balance(gInBedBillInfo_��ɽ)
    If GetErrInfo_��ɽ Then Exit Function
    
    Call Get������Ϣ(cur�ʻ�֧��)
    
    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    Call ExecuteProcedure("�����ʼ�¼�����ϴ���־")
    
    '��д���ս����¼
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gtypBalance.�ܷ��� & "," & gtypBalance.�ֽ� & "," & 0 & "," & gtypBalance.ҽ������ & "," & gtypBalance.ҽ������ & ",0," & _
        0 & "," & cur�ʻ�֧�� & ",'" & TrimTsChar(gtypBalance.˳���.Data) & "',null,null,'" & TrimTsChar(gInBedBillInfo_��ɽ.INBED_STL_BILL_NO) & "')"
    Call ExecuteProcedure("����סԺ��������")
    
    סԺ����_��ɽ = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function סԺ�������_��ɽ(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    On Error GoTo ErrHand
    
    MsgBox "��֧��סԺ����������뵽ҽ�����İ���", vbInformation, gstrSysName
    סԺ�������_��ɽ = False
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��ݱ�ʶ_��ɽ(Optional bytType As Byte, Optional lng����ID As Long) As String
'    ���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'    ������bytType-ʶ�����ͣ�0-���1-סԺ
'����:     �ջ���Ϣ��
'    ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'    2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'    3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    '��֧��סԺ
    If bytType = 1 Then
        ��ݱ�ʶ_��ɽ = frmIdentify��ɽ.GetPatient(bytType, lng����ID)
    Else
        ��ݱ�ʶ_��ɽ = ""
    End If
End Function

Private Function ��ȡ���������Ϣ(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
    Dim str��Ժ���ұ�� As String, str��Ժ�������� As String, str��Ժ������� As String
    Dim str��Ժ�������� As String, str��Ժ������� As String, str��Ժ�������� As String
    Dim strסԺҽʦ As String, strסԺ�� As String, str��Ժ��� As String
    Dim str��Ժ��� As String, str��Ժ���� As String, str��Ժ��ʽ As String
    Dim rsTemp As New ADODB.Recordset
'    ��ȡ���������Ϣ (����סԺ����||��Ժ���ұ��||��Ժ��������||��Ժ�������||��Ժ��������||��Ժ�������||סԺ��||��Ժ���||��Ժ���)
    
'    ��ȡ��Ժ�����Ϣ
    gstrSQL = "select C.���� ��Ժ���ұ��,C.���� ��Ժ��������,B.���� ��Ժ�������,B.���� ��Ժ��������, " & _
             " A.��Ժ���� ��Ժ�������,D.���� ��Ժ��������,F.��λ����,E.סԺ�� סԺ��,A.סԺҽʦ,to_char(A.��Ժ����,'yyyy-MM-dd') ��Ժ����,A.��Ժ��ʽ " & _
             " from ������ҳ A,���ű� B,���ű� C,���ű� D,������Ϣ E, " & _
             " (Select D.���� ��λ����,F.����,F.����ID,F.����ID  From ��λ�ȼ� D ,��λ״����¼ F Where F.�ȼ�ID=D.���) F " & _
             " Where A.��Ժ����ID=B.ID(+) And A.��Ժ����ID=C.ID(+) And A.��Ժ����ID=D.ID(+) And A.����ID=E.����ID ANd A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & _
             " And A.��Ժ����=F.����(+) And F.����ID(+)=A.��Ժ����ID And F.����ID(+)=A.��Ժ����ID"
    Call OpenRecordset(rsTemp, "��ȡ��Ժ�����Ϣ")
    If Not rsTemp.EOF Then
        str��Ժ���ұ�� = NVL(rsTemp!��Ժ���ұ��)
        str��Ժ�������� = NVL(rsTemp!��Ժ��������)
        str��Ժ������� = NVL(rsTemp!��Ժ�������)
        str��Ժ�������� = NVL(rsTemp!��Ժ��������)
        str��Ժ������� = NVL(rsTemp!��Ժ�������)
        str��Ժ�������� = NVL(rsTemp!��Ժ��������)
        strסԺҽʦ = NVL(rsTemp!סԺҽʦ)
        str��Ժ���� = NVL(rsTemp!��Ժ����)
        str��Ժ��ʽ = NVL(rsTemp!��Ժ��ʽ)
        strסԺ�� = NVL(rsTemp!סԺ��)
    End If
    
'    ��ȡ���Ժ��ϣ����|�������룩
    str��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, True, False, True)
    str��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, False, True)
    ��ȡ���������Ϣ = str��Ժ���ұ�� & "||" & str��Ժ�������� & "||" & _
                    str��Ժ������� & "||" & str��Ժ�������� & "||" & str��Ժ������� & "||" & _
                    str��Ժ�������� & "||" & strסԺҽʦ & "||" & strסԺ�� & "||" & str��Ժ��� & _
                    "||" & str��Ժ��� & "||" & str��Ժ���� & "||" & str��Ժ��ʽ
End Function

Private Sub ��ȡ���˻�����Ϣ(ByVal lng����ID As Long)
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select ˳��� From �����ʻ� Where ����ID=" & lng����ID & " And ����=" & TYPE_��ɽ
    Call OpenRecordset(rsTemp, "��ȡ���˵�סԺ��ˮ��")
    
    gtypBalance.˳���.Data = NVL(rsTemp!˳���)
End Sub

Private Function �Ƿ�ҽ������(ByVal lng����ID As Long) As Boolean
    Dim rsInsure As New ADODB.Recordset
    
    '��鱾���Ƿ���ҽ�������Ժ
    gstrSQL = "Select Count(*) Records From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.����ID=" & lng����ID & " And A.��ҳID=B.סԺ���� And A.����=" & TYPE_��ɽ
    Call OpenRecordset(rsInsure, "�ж��Ƿ�ҽ������")
    �Ƿ�ҽ������ = (rsInsure!Records = 1)
End Function

Private Sub Get������Ϣ(Optional ByVal cur�ʻ�֧�� As Currency = 0)
    '����Ԥ�������㷵�ص�ֵ����ʾ������Ϣ�����ڸ����ʻ��ǽӿڷ��صģ����Ʋ������޸ģ�
    With gtypBalance
'        INBED_FUND_AMT      As Single           'סԺͳ��֧�����
'        INBED_ACCT_AMT      As Single           'סԺ����֧�����
'        CASH_PAY_AMT        As Single           '�ֽ�֧�����
'        HAI_INBED_SBS_AMT   As Single           '����סԺ����֧�����
'        HAI_INBED_AMT       As Single           '����סԺ֧�����
'        HAI_INBED_REPAY_AMT As Single           '����סԺ�ٴ�֧�����
'        HAI_INBED_HIGH_AMT  As Single           '����סԺ�߶�֧�����
'        OFFICAL_HIGH_AMT    As Single           '����Ա�߶��֧�����
'        OFFICAL_INBED_AMT   As Single           '����ԱסԺ����֧�����
'        OFFICAL_ACCT_AMT    As Single           '����Ա���ʲ���֧�����
        .�����ʻ� = IIf(cur�ʻ�֧�� = 0, gInBedBillInfo_��ɽ.INBED_ACCT_AMT, cur�ʻ�֧��)
        .������� = gInBedBillInfo_��ɽ.HAI_INBED_SBS_AMT + gInBedBillInfo_��ɽ.HAI_INBED_AMT + _
        gInBedBillInfo_��ɽ.HAI_INBED_REPAY_AMT + gInBedBillInfo_��ɽ.HAI_INBED_HIGH_AMT
        .ҽ������ = gInBedBillInfo_��ɽ.INBED_FUND_AMT + gInBedBillInfo_��ɽ.OFFICAL_HIGH_AMT + _
        gInBedBillInfo_��ɽ.OFFICAL_INBED_AMT + gInBedBillInfo_��ɽ.OFFICAL_ACCT_AMT
        If cur�ʻ�֧�� <> 0 Then
            .�ֽ� = .�ܷ��� - .ҽ������ - .������� - .�����ʻ�
        End If
    End With
End Sub

Public Function �ϴ�����_��ɽ(ByVal int���� As Integer, ByVal int״̬ As Integer, ByVal str���ݺ� As String) As Boolean
    Dim intTYPE As Integer
    Dim lng����ID As Long
    Dim blnInsure As Boolean, blnUpload As Boolean, blnTrans As Boolean
    Dim rsExse As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim gcn�ϴ� As New ADODB.Connection
    On Error GoTo ErrHand
    
    gstrSQL = " Select A.ID,A.����ID,A.NO,A.���,A.��¼����,A.��¼״̬,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') �Ǽ�ʱ��,A.�շ����," & _
              " A.������ ҽ��,B.���� ��������,A.�շ�ϸĿID,D.���� ϸĿ����,C.��Ŀ���� ҽ����Ŀ����,C.ҽ������,A.ʵ�ս�� ���,A.����*Nvl(A.����,1) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�" & _
              " From ���˷��ü�¼ A,���ű� B,�շ�ϸĿ D,(Select A.*,B.���� ҽ������ From ����֧����Ŀ A,����֧������ B " & _
              "                               Where A.����=B.���� And A.����ID=B.ID And A.����=" & TYPE_��ɽ & ") C " & _
              " Where A.��¼����=" & int���� & " And A.��¼״̬=" & int״̬ & " And A.NO='" & str���ݺ� & "'" & _
              " And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And A.�շ�ϸĿID=D.ID And Nvl(A.�Ƿ��ϴ�,0)=0" & _
              " Order by A.NO,A.����ID"
    Call OpenRecordset(rsExse, "��ȡ������ϸ")
    
    With gcn�ϴ�
        If .State = 1 Then .Close
        .Open gcnOracle.ConnectionString
    End With
    
    With rsExse
        Do While Not .EOF
            If lng����ID <> !����ID Then
                '�ύ����
                If lng����ID <> 0 And blnInsure Then
                    gbytReturn_��ɽ = LS_SaveDetail(gtypBalance.˳���)
                    If GetErrInfo_��ɽ Then
                        gcn�ϴ�.RollbackTrans
                        Exit Function
                    End If
                    gcn�ϴ�.CommitTrans
                    blnTrans = False
                End If
            End If
            
            '�жϵ�ǰ�����Ƿ񱾴���ҽ����ݵǼ�
            If lng����ID <> !����ID Then blnInsure = �Ƿ�ҽ������(!����ID)
            If blnInsure Then
                If lng����ID <> !����ID Then
                    lng����ID = !����ID
                    Call ��ȡ���˻�����Ϣ(lng����ID)
                    gbytReturn_��ɽ = LS_GetInHospitalRegInfo(gtypBalance.˳���)
                    gcn�ϴ�.BeginTrans
                    blnTrans = True
                    If GetErrInfo_��ɽ Then
                        gcn�ϴ�.RollbackTrans
                        Exit Function
                    End If
                End If
                
                '�ϴ���ϸ
                intTYPE = 1
                If !ҽ������ = "����" Then intTYPE = 2
                If !ҽ������ = "����" Then intTYPE = 3
                Select Case intTYPE
                Case 1
                    gstrSQL = "select A.���,A.����,B.���� ����  " & _
                             " from ҩƷĿ¼ A,ҩƷ���� B,ҩƷ��Ϣ C " & _
                             " Where A.ҩ��ID=C.ҩ��ID And C.����=B.���� And A.ҩƷID=" & !�շ�ϸĿID
                    Call OpenRecordset(rsTemp, "��ȡҩƷ��Ϣ")
                    
                    With gDrugInfo_��ɽ
                        .INBED_REG_ID = gtypBalance.˳���.Data
                        .RECEIPT_DT = Format(rsExse!�Ǽ�ʱ��, "yyyy-MM-dd")
                        .DRUG_CATALOG_ID = rsExse!ҽ����Ŀ����
                        .DRUG_INFO = rsExse!ϸĿ����
                        .UNIT_PRC = Format(rsExse!��� / rsExse!����, "#####0.0000;-#####0.0000;0;")
                        .SRVC_CNT = rsExse!����
                        .COST_PRC = 0
                        .DRUG_TYP = NVL(rsTemp!����)
                        .DRUG_SPEC = NVL(rsTemp!���)
                        .PRODUCE_FACTORY = NVL(rsTemp!����)
                    End With
                Case 2
                    With gDiagInfo_��ɽ
                        .INBED_REG_ID = gtypBalance.˳���.Data
                        .RECEIPT_DT = Format(rsExse!�Ǽ�ʱ��, "yyyy-MM-dd")
                        .DIAG_CATALOG_ID = rsExse!ҽ����Ŀ����
                        .DIAG_ITEM_NAME = rsExse!ϸĿ����
                        .UNIT_PRC = Format(rsExse!��� / rsExse!����, "#####0.0000;-#####0.0000;0;")
                        .SRVC_CNT = rsExse!����
                    End With
                Case 3
                    With gServiceItemInfo_��ɽ
                        .INBED_REG_ID = gtypBalance.˳���.Data
                        .RECEIPT_DT = Format(rsExse!�Ǽ�ʱ��, "yyyy-MM-dd")
                        .SRVC_ITEM_ID = rsExse!ҽ����Ŀ����
                        .SRVC_NAME = rsExse!ϸĿ����
                        .UNIT_PRC = Format(rsExse!��� / rsExse!����, "#####0.0000;-#####0.0000;0;")
                        .SRVC_CNT = rsExse!����
                    End With
                End Select
                
                If Not UploadDetail(intTYPE) Then
                    gcn�ϴ�.RollbackTrans
                    Exit Function
                End If
                blnUpload = True
            End If
            .MoveNext
        Loop
        If blnUpload And blnInsure Then
            gbytReturn_��ɽ = LS_SaveDetail(gtypBalance.˳���)
            If GetErrInfo_��ɽ Then
                gcn�ϴ�.RollbackTrans
                Exit Function
            End If
            gcn�ϴ�.CommitTrans
            blnTrans = False
        End If
        
        '���ϱ��
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '�����ϴ���־����Ϊ��ϸ������ȷ�ϴ��󣬲��ܱ�֤�������ȷ��
            'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
            gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsExse("NO") & "'," & rsExse("���") & "," & rsExse("��¼����") & "," & rsExse("��¼״̬") & ")"
            Call ExecuteProcedure("�����ϴ���־")
            .MoveNext
        Loop
    End With
    
    �ϴ�����_��ɽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcn�ϴ�.RollbackTrans
End Function

Private Function UploadDetail(Optional ByVal intTYPE As Integer = 1) As Boolean
    '�ϴ�������ϸ
    'intType:1-ҩƷ;2-����;3-����
    Select Case intTYPE
    Case 1
        gbytReturn_��ɽ = LS_AddDrug(gDrugInfo_��ɽ)
    Case 2
        gbytReturn_��ɽ = LS_AddDiag(gDiagInfo_��ɽ)
    Case 3
        gbytReturn_��ɽ = LS_AddService(gServiceItemInfo_��ɽ)
    End Select
    If GetErrInfo_��ɽ Then Exit Function
    UploadDetail = True
End Function

Private Function TrimTsChar(ByVal strData As Variant) As String
    TrimTsChar = Replace(Replace(strData, " ", ""), Chr(0), "")
End Function


