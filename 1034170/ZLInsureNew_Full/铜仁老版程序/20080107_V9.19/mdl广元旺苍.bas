Attribute VB_Name = "mdl��Ԫ����"
Option Explicit
Public Enum ҵ������_��Ԫ����
    ����籣����_���� = 0
    ��òα���Ա����_����
    ��ȡ�ʻ����_����
    ���κ�����_����
    �����κ�����_����
    �Ͽ��κ�����_����
    �����ʻ�����_����
    �����ʻ�����_���_����
    ���ѳ���_����
'    ��ӡ��Ժ���㱨����
'    ��ӡסԺ��Ա������㵥
'    ��ȡסԺ��¼��
    ��ȡҩƷ��Ϣ
End Enum
Private Type InitbaseInfor
    ģ������ As Boolean                     '��ǰ�Ƿ���ģ���ȡҽ���ӿ�����
    ҽԺ���� As String                      '��ʼҽԺ����
    �������� As String                      'Ĭ�ϵ��籣��������
    
End Type
Public InitInfor_��Ԫ���� As InitbaseInfor

Private Type �������
    ҽ������    As String
    ҽ��֤��    As String
    ���֤����  As String
    ��¼��      As String
    ����        As String
    �Ա�        As String
    ��������    As String
    ����        As Integer
    ��λ����    As String
    ��������    As String
    
    �ʻ����    As String
    �����ܶ�    As Double
    ����        As String
    �籣����    As Long
    ����id      As Long
End Type

Private Type ��������
    ���� As String
    ����    As String
    ����ǰ�ʻ���� As Double
    �����ʻ�֧����� As Double
    �Էѽ�� As Double
    ���Ѻ��ʻ���� As Double
    ����ʱ��  As String
    ǰ�˵��ݺ�  As String
    ���ĵ��ݺ�  As String
    ������  As String
    ����Ա����  As String
    ǰ������  As String
    
    ����id As Long
    �����־ As Byte    '0-����,1-סԺ
End Type
Private g�������� As ��������
Public g�������_��Ԫ���� As �������
Public gcnOracle_��Ԫ���� As ADODB.Connection     '�м������

Private gbln������� As Boolean
Private gbln�Ѿ���ʼ As Boolean             '�Ѿ�����ʼ����.

'1.����籣����_���Ա�ź������б�
Private Declare Function GetSBJGLB Lib "cdgk_Yb.dll" Alias "GETSBJGLB" () As String
'===============================================================================================================
'ԭ��:FUNCTION GETSBJGLB:PCHAR
'����: ����籣����_���Ա�ź������б�
'��ڲ���: ��
'���ڲ���: ��
'����: A�籣�������+�м����+A�籣��������+�м����+B�籣�������+�м����+B�籣��������+����
'===============================================================================================================

'2����òα���Ա�Ļ�������
Private Declare Function GETKZL Lib "cdgk_Yb.dll" () As String
'===============================================================================================================
'ԭ��:FUNCTION GETKZL:PCHAR
'����: ��òα���Ա�Ļ�������
'��ڲ���:
'���ڲ���: ��
'����: OK(�������Ϣ)@$ҽ������||ҽ��֤��||���˼�¼��||����||���֤����||��λ����||�Ա�||�������ڣ���ʽ��YYYY-MM-DD��
'===============================================================================================================

'3.�����ʻ�����ѯ
Private Declare Function GETZHYE Lib "cdgk_Yb.dll" (ByVal str�������� As String, ByVal strPassWord As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GETZHYE(YBJGBH,CPASSWORD:PCHAR):PCHAR
'����: ��óֿ���Ա�����ʻ����
'��ڲ���:YBJGBH  PCHAR   ���ջ������
'         CPASSWORD   PCHAR   �ֿ��˿�����
'���ڲ���: ��
'����:  OK(�������Ϣ)@$�����ʻ����
'===============================================================================================================

'4.���κ������Ƿ����ӳɹ�
Private Declare Function CheckCon Lib "cdgk_Yb.dll" () As String
'===============================================================================================================
'ԭ��:FUNCTION CHECKCON:PCHAR;
'����:���κ������Ƿ����ӳɹ�
'��ڲ���:
'����:OK�������Ϣ
'===============================================================================================================

'5.�����κ�����
Private Declare Function RasDial Lib "cdgk_Yb.dll" (ByVal str�������� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION RASDIAL(SBXJGBH:PCHAR):PCHAR
'����:�κ���ѡ����籣�֣����佨������
'��ڲ���:SBXJGBH PCHAR   ���ջ������
'����:  �ɹ�    ������HIS�κ���״̬����ʾ"����"
'       ʧ�� ������Ϣ
'===============================================================================================================



'6.�Ͽ����籣�ֵ�����
Private Declare Function DisDial Lib "cdgk_Yb.dll" () As String
'===============================================================================================================
'ԭ��:FUNCTION DISDIAL:PCHAR
'����:�κ���ѡ����籣�֣����佨������
'��ڲ���:
'����:
'===============================================================================================================

'7.�����ʻ�����
Private Declare Function GRZHXF_CF Lib "cdgk_Yb.dll" (ByVal str������� As String, str������ As String, _
            ByVal str��ϸ���� As String, ByVal strPassWord As String, ByVal str����Ա As String) As String
'===============================================================================================================
'ԭ��:Function GRZHXF_CF()(YBJGBH,CFH:PCHAR;CFMXDATA:PCHAR; CPASSWORD:PCHAR;CCZYXM:PCHAR):PCHAR
'����:���и����ʻ�����
'��ڲ���:YBJGBH  PCHAR   ���ջ������
'        CFH PCHAR   ������
'        CFMXDATA    PCHAR   ������ϸ����    ��ʽ˵��������1(ҽ��ҩƷ���+�м����+����+�м��������+)+�м����+        ����        ����N(ҽ��ҩƷ���+�м����+����+�м����+����
'        CPASSWORD   PCHAR   �ֿ��˿�����
'        CCZYXM  PCHAR   ����Ա����
'����:�����ʻ�������Ϣ(OK@$�����ʻ�������Ϣ)
'   ��ʽ:����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
'===============================================================================================================


'8.�����ʻ����ѣ�ֱ���������ѽ�

Private Declare Function GRZHXF_JE Lib "cdgk_Yb.dll" (ByVal str������� As String, str������ As String, _
             ByVal strPassWord As String, ByVal str����Ա As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GRZHXF_JE(YBJGBH,XFJE:PCHAR; CPASSWORD:PCHAR;CCZYXM:PCHAR):PCHAR
'����:���и����ʻ�����
'��ڲ���:YBJGBH  PCHAR   ���ջ������
'    XFJE    PCHAR   ���ѽ��(��֤ΪС�������ұ�����λС��)
'    CPASSWORD   PCHAR   �ֿ��˿�����
'    CCZYXM  PCHAR   ����Ա����
'����:�����ʻ�������Ϣ(OK@$�����ʻ�������Ϣ)
'   ��ʽ:����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
'===============================================================================================================

'9.���ѳ���

Private Declare Function XFCZ Lib "cdgk_Yb.dll" (ByVal str������� As String, str���ĵ��ݺ� As String, _
             ByVal strPassWord As String, ByVal str����Ա As String) As String
'===============================================================================================================
'ԭ��:FUNCTION XFCZ(YBJGBH ��CZXDJH:PCHAR; CPASSWORD:PCHAR;CCZYXM:PCHAR):PCHAR
'����:���Ѿ����ѵļ�¼���г�����
'��ڲ���:YBJGBH  PCHAR   ���ջ������
'        cZXDJH  PCHAR   ���ĵ��ݺ�(����ʱ����)
'        CPASSWORD   PCHAR   �ֿ��˿�����
'        CCZYXM  PCHAR   ����Ա����
'����:�����ʻ�������Ϣ(OK@$�����ʻ�������Ϣ)
'   ��ʽ:����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
'===============================================================================================================












'12.��ӡ��Ժ���㱨����
Private Declare Function JSReport Lib "cdgk_Yb.dll" (ByVal str��ʼסԺ�� As String, ByVal str����סԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION JSREPORT(ASTARTZYH,AENDZYH:PCHAR):PCHAR;STDCALL
'����:��ӡ�籣�����ṩ�Ķ�̬����Ŀǰ�����������ö�̬����"סԺ����ͳ�Ʊ����䣩"��"����סԺ���㵥"��"סԺ����ͳ�Ʊ�"���ű���ʹ��"21����ȡ��������"�������Զ����±��ر���
'��ڲ���:
'    ASTARTZYH   PCHAR   ��ӡ��ʼסԺ��
'    AENDZYH PCHAR      ��ӡ����סԺ��
'   ע��:
'    1 ?����סԺ��֮�����е�סԺ��¼����Ϊͬһ���籣��?
'    2����ֻ��ӡһ��סԺ�ŵı���ʱ����������ֵһ����
'���ڲ���: ��
'����:����ע�ⷵ��ֵ
'===============================================================================================================

'13.��ӡסԺ��Ա������㵥
Private Declare Function CWJSReport Lib "cdgk_Yb.dll" (ByVal str��ʼסԺ�� As String, ByVal str����סԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION CWJSREPORT(ASTARTZYH,AENDZYH:PCHAR):PCHAR;STDCALL;
'����:��ӡסԺ��Ա������㵥��
'��ڲ���:
'    ASTARTZYH   PCHAR   ��ӡ��ʼסԺ��
'    AENDZYH PCHAR      ��ӡ����סԺ��
'   ע��:
'    1 ?����סԺ��֮�����е�סԺ��¼����Ϊͬһ���籣��?
'    2����ֻ��ӡһ��סԺ�ŵı���ʱ����������ֵһ����
'���ڲ���: ��
'����:����ע�ⷵ��ֵ
'===============================================================================================================

'14.��ȡ��������
Private Declare Function GetJCXX Lib "cdgk_Yb.dll" (ByVal str������� As String, ByVal str���ر�־ As String) As String
'===============================================================================================================
'ԭ��:GETJCXX(SBXJGBH:PCHAR;DOWNALL:INTEGER):PCHAR
'����:��ָ�����籣������ȡ��������
'��ڲ���:
'    SBXJGBH PCHAR   ���ջ������
'    DOWNALL PCHAR   ��ֵΪ0ʱ��ʾ���ر���ҽ�����ݿ���û�еĻ������ϣ�Ϊ����ʱ��ʾȫ����������
'���ڲ���: ��
'����:'OK'�������Ϣ
'===============================================================================================================

'15 ����סԺ�ŵõ�סԺ��¼��
Private Declare Function GetZYIDByZyBH Lib "cdgk_Yb.dll" (ByVal strסԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GETZYIDBYZYBH(AZYH:PCHAR):PCHAR
'����:����סԺ�ŵõ�סԺ��¼��
'��ڲ���:
'   AZYH    PCHAR   סԺ��'���ڲ���: ��
'����:'OK'@$סԺ��¼�Ż������Ϣ
'===============================================================================================================


'19 ����ҩƷ��ŵõ�ҩƷ��Ϣ
Private Declare Function GetSINYPXX Lib "cdgk_Yb.dll" (ByVal str�������� As String, ByVal strҩƷ���� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GETSINYPXX(SBXJGBH,CYPBH:PCHAR):PCHAR
'����:����ҩƷ��ŵõ�ҩƷ��Ϣ
'��ڲ���:
'    SBXJGBH PCHAR   ���ջ������
'    CYPBH   PCHAR   ҩƷ���
'����:OK@$���:ҩƷ||��������:��Ī�����ƣ�����ά��أ�||������λ:֧||��������:0||�Էѱ���:20
'===============================================================================================================




Public Function ҽ����ʼ��_��Ԫ����() As Boolean
    Dim strReg As String, strOutPut As String
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    '��ʼģ��ӿ�
    Call GetRegInFor(g����ģ��, "����", "ģ��ӿ�", strReg)
    If Val(strReg) = 1 Then
        InitInfor_��Ԫ����.ģ������ = True
    Else
        InitInfor_��Ԫ����.ģ������ = False
    End If
    
   Call GetRegInFor(g����ģ��, "ҽ��", "�籣��������", strReg)
   
   InitInfor_��Ԫ����.�������� = strReg
   g�������_��Ԫ����.�������� = strReg
   
   If strReg = "" Then
        MsgBox "��δ����Ĭ�ϵ��籣�������룬�����������!"
        Exit Function
   End If
   
    'ȡҽԺ����
    gstrSQL = "Select ҽԺ���� From ������� Where ���=" & TYPE_��Ԫ����
    Call OpenRecordset(rsTemp, "��ȡҽԺ����")
    InitInfor_��Ԫ����.ҽԺ���� = Nvl(rsTemp!ҽԺ����)
    
    
    '�м������
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=" & TYPE_��Ԫ����
    Call OpenRecordset(rsTemp, "�山ҽ��")
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "ҽ���û���"
                strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ��������"
                strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ���û�����"
                strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "��鲦������"
                 gbln������� = Nvl(rsTemp("����ֵ"), 0) = 1
        End Select
        rsTemp.MoveNext
    Loop
    
    Set gcnOracle_��Ԫ���� = New ADODB.Connection
    If OraDataOpen(gcnOracle_��Ԫ����, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ�ҽ���м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
    
   '�����κ�����
   If gbln�Ѿ���ʼ = False And gbln������� Then
       If ������������() = False Then Exit Function
   End If
   
   If gbln������� Then
        '���κ�����
        If ҵ������_��Ԫ����(���κ�����, "", strOutPut) = False Then
             Exit Function
        End If
    End If
    gbln�Ѿ���ʼ = True
    ҽ����ʼ��_��Ԫ���� = True
End Function

Public Function ҽ����ֹ_��Ԫ����() As Boolean
    Dim strOutPut As String
    
    If gcnOracle_��Ԫ����.State = 1 Then
        gcnOracle_��Ԫ����.Close
    End If
    '�����κ�����
   Call ҵ������_��Ԫ����(�Ͽ��κ�����, "", strOutPut)
    Err = 0
    On Error Resume Next
    ҽ����ֹ_��Ԫ���� = True
End Function

Public Function ��ݱ�ʶ_��Ԫ����(Optional bytType As Byte, Optional lng����ID As Long) As String
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    '���أ��ջ���Ϣ��
    Err = 0
    On Error GoTo ErrHand:
    If bytType = 0 Or bytType = 3 Then Exit Function
    
    ��ݱ�ʶ_��Ԫ���� = frmIdentify��Ԫ����.GetPatient(bytType, lng����ID)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_��Ԫ���� = ""
End Function


Public Function �������_��Ԫ����(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(�ʻ����,0) as �ʻ���� from �����ʻ� where ����ID='" & lng����ID & "' and ����=" & TYPE_��Ԫ����
    Call OpenRecordset(rsTemp, "��ȡ�����ʻ����")
    
    If rsTemp.EOF Then
        �������_��Ԫ���� = 0
    Else
        �������_��Ԫ���� = rsTemp("�ʻ����")
    End If
End Function
Public Function �����������_��Ԫ����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    �����������_��Ԫ���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ������������() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ҵ�����ҵ������
    '--�����:strinPutString-���봮,������˳��,��tab���ָ��Ĵ��봮
    '--������:strOutPutString-�����,������˳��,��tab���ָ��ķ��ش�
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Static str������� As String
    Dim strInput As String, strOutPut As String
    ������������ = False
    
    Err = 0: On Error GoTo ErrHand:
    If str������� <> g�������_��Ԫ����.�������� Then
        '��������Ƿ���������
        If str������� = "" Then
            '�����һ��Զ��,��Ͽ�
            If ҵ������_��Ԫ����(�����κ�����_����, g�������_��Ԫ����.��������, strOutPut) = False Then
                Exit Function
            End If
        Else
            '��ʾ�������������ϵĲ���,����Ͽ�����
            Call ҵ������_��Ԫ����(�Ͽ��κ�����_����, "", strOutPut)
            If ҵ������_��Ԫ����(�����κ�����_����, g�������_��Ԫ����.��������, strOutPut) = False Then Exit Function
        End If
        If ҵ������_��Ԫ����(���κ�����_����, "", strOutPut) = False Then Exit Function
    Else
        If ҵ������_��Ԫ����(���κ�����_����, "", strOutPut) = False Then
            '�����½�����������
            If ҵ������_��Ԫ����(�����κ�����_����, g�������_��Ԫ����.��������, strOutPut) = False Then
                Exit Function
            End If
        End If
    End If
    str������� = g�������_��Ԫ����.��������
    ������������ = True
    Exit Function
ErrHand:
        If ErrCenter = 1 Then Resume
End Function
Public Function �������_��Ԫ����(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
        '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim strInput As String, strOutPut As String
    Dim lng����ID  As Long
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strArr As Variant
    If ������������() = False Then Exit Function
    
    On Error GoTo errHandle
    
    Call DebugTool("�����������")
    
    gstrSQL = "" & _
        "   Select a.*,a.*,a.����*a.���� as ����,a.ʵ�ս��/(nvl(a.����,1)*nvl(a.����,1)) as ���� " & _
        "   From ���˷��ü�¼ a " & _
        "   Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
        
    Call OpenRecordset(rs��ϸ, "��ȡ��ϸ��¼")
    
    If rs��ϸ.EOF = True Then
        MsgBox "û����д�շѼ�¼", vbExclamation, gstrSysName
        Exit Function
    End If

    lng����ID = rs��ϸ("����ID")
    
    If g�������_��Ԫ����.����id <> lng����ID Then
        MsgBox "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    g��������.����id = lng����ID
    g��������.�����־ = 0
    'д����ϸ
    If ������ϸд��(rs��ϸ, False) = False Then Exit Function
    
    '��ʾ��ᴦ��ʽ
    If ���㷽ʽ����() = False Then
        Exit Function
    End If
    
    
    
    Dim dbl�����ʻ� As Double
    dbl�����ʻ� = ��ȡ�����ʻ�֧��()
    If dbl�����ʻ� <> g��������.�����ʻ�֧����� Then
        '���¸����ʻ�֧��
        '��:YBJGBH  PCHAR   ���ջ������
        '    XFJE    PCHAR   ���ѽ��(��֤ΪС�������ұ�����λС��)
        '    CPASSWORD   PCHAR   �ֿ��˿�����
        '    CCZYXM  PCHAR   ����Ա����
        '����:����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
        strInput = g�������_��Ԫ����.��������
        strInput = strInput & vbTab & Format(dbl�����ʻ�, "###0.00;-###0.00;0.00;0.00")
        strInput = strInput & vbTab & g�������_��Ԫ����.����
        strInput = strInput & vbTab & gstrUserName
        If ҵ������_��Ԫ����(�����ʻ�����_���_����, strInput, strOutPut) = False Then Exit Function
        If strOutPut = "" Then Exit Function
        strArr = Split(strOutPut, "||")
        
        With g��������
            .���� = strArr(0)
            .���� = strArr(1)
            .����ǰ�ʻ���� = Val(strArr(2))
            .�����ʻ�֧����� = Val(strArr(3))
            .�Էѽ�� = Val(strArr(4))
            .���Ѻ��ʻ���� = Val(strArr(5))
            .����ʱ�� = strArr(6)
            .ǰ�˵��ݺ� = strArr(7)
            .���ĵ��ݺ� = strArr(8)
            .������ = strArr(9)
            .����Ա���� = strArr(10)
            .ǰ������ = strArr(11)
        End With
    End If
       
    '��д�����
    Call DebugTool("��д�����¼")
    

    
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(����ǰ�ʻ����),�ۼ�ͳ�ﱨ��_IN(���Ѻ��ʻ����),סԺ����_IN(��),����(��),�ⶥ��_IN(��),ʵ������_IN(��),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(��),�����Ը����_IN(�Էѽ��),
    '   ����ͳ����_IN(��),ͳ�ﱨ�����_IN(��),���Ը����_IN(��),�����Ը����_IN(��),�����ʻ�֧��_IN(�����ʻ�֧�����),"
    '   ֧��˳���_IN(���ĵ��ݺ�),��ҳID_IN(��),��;����_IN,��ע_IN(ǰ�˵��ݺ�|������|����Ա����|ǰ������)
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
   
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL," & g��������.����ǰ�ʻ���� & "," & g��������.���Ѻ��ʻ���� & ",null,0,0,0," & _
            g�������_��Ԫ����.�����ܶ� & ",0," & g��������.�Էѽ�� & "," & _
          "0,0,0,0," & g��������.�����ʻ�֧����� & ",'" & _
            g��������.���ĵ��ݺ� & " ',NULL,NULL,'" & g��������.ǰ�˵��ݺ� & "|" & g��������.������ & "|" & g��������.����Ա���� & "|" & g��������.ǰ�˵��ݺ� & "')"
            
    Call ExecuteProcedure("��������¼")
    '---------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------
    �������_��Ԫ���� = True
    Exit Function

Err������:

'��ڲ���:YBJGBH  PCHAR   ���ջ������
'        cZXDJH  PCHAR   ���ĵ��ݺ�(����ʱ����)
'        CPASSWORD   PCHAR   �ֿ��˿�����
'        CCZYXM  PCHAR   ����Ա����
    strInput = g�������_��Ԫ����.��������
    strInput = strInput & vbTab & g��������.���ĵ��ݺ�
    strInput = strInput & vbTab & g�������_��Ԫ����.����
    strInput = strInput & vbTab & gstrUserName
    
    If ҵ������_��Ԫ����(���ѳ���_����, strInput, strOutPut) = False Then Exit Function
'����:�����ʻ�������Ϣ(OK@$�����ʻ�������Ϣ)
'   ��ʽ:����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
    If strOutPut = "" Then Exit Function
     strArr = Split(strOutPut, "||")
    
    With g��������
        .���� = strArr(0)
        .���� = strArr(1)
        .����ǰ�ʻ���� = Val(strArr(2))
        .�����ʻ�֧����� = Val(strArr(3))
        .�Էѽ�� = Val(strArr(4))
        .���Ѻ��ʻ���� = Val(strArr(5))
        .����ʱ�� = strArr(6)
        .ǰ�˵��ݺ� = strArr(7)
        .���ĵ��ݺ� = strArr(8)
        .������ = strArr(9)
        .����Ա���� = strArr(10)
        .ǰ������ = strArr(11)
    End With

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function ����������_��Ԫ����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    
    Dim intMouse As Integer
    Dim lng����ID  As Long
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutPut As String
    Dim strArr As Variant
    
    ����������_��Ԫ���� = False
    
    '�����֤
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    If ��ݱ�ʶ_����(0, lng����ID) = "" Then
        Screen.MousePointer = intMouse
        Exit Function
    End If
    Screen.MousePointer = intMouse
    
    gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "����ҽ��")
    lng����ID = rsTemp("����ID")
    
    
    
    gstrSQL = "Select * From ���˷��ü�¼ " & _
        " Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
        
    Call OpenRecordset(rs��ϸ, "��ȡ������¼")
    g�������_��Ԫ����.�����ܶ� = 0
    With rs��ϸ
        Do While Not .EOF
                'д�ϴ���־
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,null)"
            ExecuteProcedure "�����ϴ���־"
            g�������_��Ԫ����.�����ܶ� = g�������_��Ԫ����.�����ܶ� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
    End With
    
    '����:
    gstrSQL = "Select ֧��˳��� from ���ս����¼ where ����=1 and ��¼id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���ĵ��ݺ�"
    If rsTemp.EOF Then
        ShowMsgbox "�����ڽ����¼,���ܳ���!"
        Exit Function
    End If
    
    '��ڲ���:YBJGBH  PCHAR   ���ջ������
    '        cZXDJH  PCHAR   ���ĵ��ݺ�(����ʱ����)
    '        CPASSWORD   PCHAR   �ֿ��˿�����
    '        CCZYXM  PCHAR   ����Ա����
    strInput = g�������_��Ԫ����.��������
    strInput = strInput & vbTab & Nvl(rsTemp!֧��˳���)
    strInput = strInput & vbTab & g�������_��Ԫ����.����
    strInput = strInput & vbTab & gstrUserName
    
    If ҵ������_��Ԫ����(���ѳ���_����, strInput, strOutPut) = False Then Exit Function
    '����:�����ʻ�������Ϣ(OK@$�����ʻ�������Ϣ)
    '   ��ʽ:����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
    If strOutPut = "" Then Exit Function
     strArr = Split(strOutPut, "||")
    
    With g��������
        .���� = strArr(0)
        .���� = strArr(1)
        .����ǰ�ʻ���� = Val(strArr(2))
        .�����ʻ�֧����� = Val(strArr(3))
        .�Էѽ�� = Val(strArr(4))
        .���Ѻ��ʻ���� = Val(strArr(5))
        .����ʱ�� = strArr(6)
        .ǰ�˵��ݺ� = strArr(7)
        .���ĵ��ݺ� = strArr(8)
        .������ = strArr(9)
        .����Ա���� = strArr(10)
        .ǰ������ = strArr(11)
    End With
    ����������_��Ԫ���� = True
        
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(����ǰ�ʻ����),�ۼ�ͳ�ﱨ��_IN(���Ѻ��ʻ����),סԺ����_IN(��),����(��),�ⶥ��_IN(��),ʵ������_IN(��),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(��),�����Ը����_IN(�Էѽ��),
    '   ����ͳ����_IN(��),ͳ�ﱨ�����_IN(��),���Ը����_IN(��),�����Ը����_IN(��),�����ʻ�֧��_IN(�����ʻ�֧�����),"
    '   ֧��˳���_IN(���ĵ��ݺ�),��ҳID_IN(��),��;����_IN,��ע_IN(ǰ�˵��ݺ�|������|����Ա����|ǰ������)
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
   
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL," & -1 * g��������.����ǰ�ʻ���� & "," & -1 * g��������.���Ѻ��ʻ���� & ",null,0,0,0," & _
           -1 * g�������_��Ԫ����.�����ܶ� & ",0," & -1 * g��������.�Էѽ�� & "," & _
          "0,0,0,0," & -1 * g��������.�����ʻ�֧����� & ",'" & _
            g��������.���ĵ��ݺ� & " ',NULL,NULL,'" & g��������.ǰ�˵��ݺ� & "|" & g��������.������ & "|" & g��������.����Ա���� & "|" & g��������.ǰ�˵��ݺ� & "')"
            
    Call ExecuteProcedure("��������¼")
    '---------------------------------------------------------------------------------------------
    ����������_��Ԫ���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_��Ԫ����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    '�����˵�״̬�����޸�
    ShowMsgbox "��ҽ���ӿڲ�֧��סԺ����"
    
    ��Ժ�Ǽ�_��Ԫ���� = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_��Ԫ���� = False
End Function
Private Function Get���״���(ByVal intType As ҵ������_��Ԫ����, Optional bln������ As Boolean = False) As String
    '������û��
    Select Case intType
        Case ����籣����_����
            Get���״��� = IIf(bln������, "����籣����", "01")
        Case ��òα���Ա����_����
            Get���״��� = IIf(bln������, "��òα���Ա����", "02")
        Case ��ȡ�ʻ����_����
                Get���״��� = IIf(bln������, "��ȡ�ʻ����", "03")
        Case ���κ�����_����
            Get���״��� = IIf(bln������, "���κ�����", "04")
        Case �����κ�����_����
            Get���״��� = IIf(bln������, "�����κ�����", "05")
        Case �Ͽ��κ�����_����
            Get���״��� = IIf(bln������, "�Ͽ��κ�����", "06")
        Case �����ʻ�����_����
            Get���״��� = IIf(bln������, "�����ʻ�����", "07")
        Case �����ʻ�����_���_����
            Get���״��� = IIf(bln������, "�����ʻ�����_���", "08")
        Case ���ѳ���_����
            Get���״��� = IIf(bln������, "���ѳ���", "09")
        Case Else
            Get���״��� = IIf(bln������, "����Ľ��״���", "-1")
    End Select
End Function
Public Function ҵ������_��Ԫ����(ByVal intType As ҵ������_��Ԫ����, strInputString As String, strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ҵ�����ҵ������
    '--�����:strinPutString-���봮,������˳��,��tab���ָ��Ĵ��봮
    '--������:strOutPutString-�����,������˳��,��tab���ָ��ķ��ش�
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strInput As String, lngReturn As Long, strOutPut As String, strReturn As String
    Dim strInValue(0 To 20) As String
    
    Dim str���״��� As String
    Dim i As Integer
    Dim strArr
    
    str���״��� = Get���״���(intType, True)
    strInput = strInputString
    DebugTool "����ҵ��������(ҵ�����ʹ���Ϊ:" & intType & " ҵ�����ƣ�" & str���״��� & ")" & vbCrLf & "        �������Ϊ:" & strInputString
    
    ҵ������_��Ԫ���� = False
    If InitInfor_��Ԫ����.ģ������ Then
        '��ȡģ������
        Readģ������ intType, strInput, strOutPutstring
         ҵ������_��Ԫ���� = True
        Exit Function
    End If
    strArr = Split(strInputString, vbTab)
    For i = 0 To UBound(strArr)
        strInValue(i) = strArr(i)
    Next
    
    Err = 0
    On Error GoTo ErrHand:
    
    Select Case intType
        Case ����籣����_����
            strOutPut = GetSBJGLB()
            
            If strOutPut = "" Then
                MsgBox "��ȡ�籣����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = strArr(1)
        Case ��òα���Ա����_����
            strOutPut = GETKZL()
            If strOutPut = "" Then
                MsgBox "��òα���Ա����_����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = strArr(1)
        Case ��ȡ�ʻ����_����
            strOutPut = GETZHYE(strInValue(0), strInValue(1))
            ''OK'+�м����+�����ʻ����
            If strOutPut = "" Then
                MsgBox "��ȡ�ʻ����_ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = strArr(1)
        
        Case ���κ�����_����
            strOutPut = CheckCon()
            If strOutPut = "" Then
                MsgBox "���κ�����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = ""
        Case �����κ�����_����
            strOutPut = RasDial(strInValue(0))
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = ""
        Case �Ͽ��κ�����_����
            strOutPut = DisDial()
            strOutPut = ""
        Case �����ʻ�����_����
            strOutPut = GRZHXF_CF(strInValue(0), strInValue(1), strInValue(2), strInValue(3), strInValue(4))
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            '����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
            strOutPut = strArr(1)
        Case �����ʻ�����_���_����
            strOutPut = GRZHXF_JE(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            '����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
            strOutPut = strArr(1)
        Case ���ѳ���_����
            strOutPut = XFCZ(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            '����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
            strOutPut = strArr(1)
        
'        Case ��Ժ�Ǽ�
'            '
'            strOutput = RYDJ(strInValue(0), Replace(strInValue(1), vbTab & "|", "||"), strInValue(2))
'            If strOutput = "" Then
'                MsgBox "��Ժ�Ǽ�ʱ,�����˿�ֵ��", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = strArr(1)
'        Case ȡ����Ժ�Ǽ�
'            strOutput = ZYQX(strInValue(0))
'            If strOutput = "" Then
'                MsgBox "ȡ����Ժ�Ǽ�ʱ,�����˿�ֵ��", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
'        Case ��Ժ�Ǽ�
'            strOutput = CYCS(strInValue(0), strInValue(1))
'            If strOutput = "" Then
'                MsgBox "��Ժ�Ǽ�ʱ,�����˿�ֵ��", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
'        Case ȡ����Ժ�Ǽ�
'            strOutput = CYCSQX(strInValue(0))
'            If strOutput = "" Then
'                MsgBox "ȡ����Ժ�Ǽ�ʱ,�����˿�ֵ��", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
'        Case ���Ӵ�������
'            strOutput = AddCFJL(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
'            If strOutput = "" Then
'                MsgBox "���Ӵ�������ʱ,�����˿�ֵ��", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = strArr(1)
'        Case ���Ӵ�����ϸ
'            strOutput = AddCFMX(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
'            If strOutput = "" Then
'                MsgBox "���Ӵ�����ϸʱ,�����˿�ֵ��", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
'            For i = 1 To UBound(strArr)
'                strOutput = "||" & strArr(i)
'            Next
'            If strOutput <> "" Then
'                strOutput = Mid(strOutput, 3)
'            End If
'        Case ɾ���������ݼ�����ϸ
'            strOutput = DELCFJL(strInValue(0))
'            If strOutput = "" Then
'                MsgBox "ɾ���������ݼ�����ϸʱ,�����˿�ֵ��", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
'        Case ������������
'            strOutput = CFCS(strInValue(0), strInValue(1))
'            If strOutput = "" Then
'                MsgBox "������������ʱ,�����˿�ֵ��", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
'        Case ��Ժ����
'            strOutput = CFCS(strInValue(0), strInValue(1))
'            If strOutput = "" Then
'                MsgBox "��Ժ����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = strArr(1)
'        Case ȡ����Ժ����
'            strOutput = CYJSQX(strInValue(0))
'            If strOutput = "" Then
'                MsgBox "��Ժ����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
'        Case ��ӡ��Ժ���㱨����
'            strOutput = JSReport(strInValue(0), strInValue(1))
'            strOutput = ""
'        Case ��ӡסԺ��Ա������㵥
'            strOutput = CWJSReport(strInValue(0), strInValue(1))
'            strOutput = ""
        
        Case ������Ա��������
            '�������ظ���"��ӡסԺ��Ա������㵥"
'            strOutPut = CWJSREPORT(strInValue(0))
'              If strOutPut = "" Then
'                MsgBox "������Ա��������ʱ,�����˿�ֵ��", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutPut, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutPut = ""
        Case ��ȡ��������
        
            strOutPut = GetJCXX(strInValue(0), strInValue(1))
              If strOutPut = "" Then
                MsgBox "��ȡ��������ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = ""
        Case ��ȡסԺ��¼��
            strOutPut = GetZYIDByZyBH(strInValue(0))
            If strOutPut = "" Then
                MsgBox "��ȡסԺ��¼��ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = strArr(1)
        Case ��ȡҩƷ��Ϣ
             strOutPut = GetSINYPXX(strInValue(0), strInValue(1))
            If strOutPut = "" Then
                MsgBox "��ȡҩƷ��Ϣʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = strArr(1)
    End Select
    strOutPutstring = strOutPut
    ҵ������_��Ԫ���� = True
    DebugTool "ҵ������ɹ�(ҵ������Ϊ:" & intType & ")." & vbCrLf & "�������Ϊ" & vbCrLf & strInputString & vbCrLf & "�������Ϊ:" & vbCrLf & strReturn
     Exit Function
    
ErrHand:
    DebugTool "ҵ������ʧ��(ҵ������Ϊ:" & intType & ")." & vbCrLf & "�������Ϊ" & vbCrLf & strInputString & vbCrLf & "�������Ϊ:" & vbCrLf & strReturn
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_��Ԫ����(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
            
    '���˺�:20040923���ӵ�
    ��Ժ�Ǽǳ���_��Ԫ���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_��Ԫ����(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutPut As String
    
    Err = 0
    On Error GoTo ErrHand:
    
    ��Ժ�Ǽ�_��Ԫ���� = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_��Ԫ���� = False
End Function
Public Function ��Ժ�Ǽǳ���_��Ԫ����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '��Ժ�Ǽǳ���
    ��Ժ�Ǽǳ���_��Ԫ���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_��Ԫ����(lng����ID As Long, ByVal lng����ID As Long) As Boolean
  '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    סԺ����_��Ԫ���� = True
    Exit Function
End Function
Public Function סԺ�������_��Ԫ����(lng����ID As Long) As Boolean
     '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    סԺ�������_��Ԫ���� = True
End Function
Public Function �����Ǽ�_��Ԫ����(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ�������ϸ����
    '--�����:
    '--������:
    '--��  ��:�ϴ��ɹ�����True,����False
    '-----------------------------------------------------------------------------------------------------------


    �����Ǽ�_��Ԫ���� = True
End Function

Private Function Readģ������(ByVal intҵ������ As ҵ������_��Ԫ����, ByVal strInputString As String, ByRef strOutPutstring As String)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--��  ��:ͨ���ù��ܶ�ȡģ������,�Ա����
    '--�����:
    '--������:
    '--��  ��:�ִ�
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    Dim str As String
    Dim strName As String
    
    If intҵ������ = ��ȡ�������� Then
        strFile = App.Path & "\������.txt"
    Else
        strFile = App.Path & "\ģ���ύ��.txt"
    End If
    
    
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    strName = Get���״���(intҵ������, True)
    
    Dim blnStart As Boolean
    Dim strArr
    
    Err = 0
    On Error GoTo ErrHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            blnStart = False
            str = ""
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                If intҵ������ = ��ȡ�������� Then
                    strArr = Split(strText, vbTab & "|")
                    If Val(strArr(0)) = 1 Then
                            str = strArr(1)
                            Exit Do
                    End If
                Else
                        If blnStart Then
                            If strText = "" Then
                                strText = "" & vbTab & "|"
                            End If
                            strArr = Split(strText, vbTab & "|")
                            
                            If Val(strArr(0)) = 1 Then
                                str = strArr(1)
                                Exit Do
                            End If
                        Else
                             If "<" & strName & ">" = strText Then
                                 blnStart = True
                             End If
                        End If
                        If "</" & strName & ">" = strText Then
                            Exit Do
                        End If
                End If
            Loop
            objText.Close
            strOutPutstring = str
    End If
'    If InStr(1, strOutPutstring, "@$") <> 0 Then
'        strOutPutstring = Split(strOutPutstring, "@$")(1)
'    End If
    Exit Function
ErrHand:
    DebugTool Err.Description
    Exit Function
End Function
Private Sub OpenRecordset_��Ԫ����(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSql As String = "")
'���ܣ��򿪼�¼��
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSql = "", gstrSQL, strSql))
    rsTemp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle_��Ԫ����, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Public Function סԺ�������_��Ԫ����(rsExse As Recordset, ByVal lng����ID As Long, Optional bln���ʴ� As Boolean = True) As String
    
    'rsExse:�ַ���
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    ShowMsgbox "��ҽ���ӿڲ�֧��סԺ����"
    Exit Function
End Function
Public Function ҽ������_��Ԫ����(ByVal lng���� As Long, ByVal lngҽ������ As Integer) As Boolean
    ҽ������_��Ԫ���� = frmSet��Ԫ����.��������
End Function
Public Sub ExecuteProcedure_��Ԫ����(ByVal strCaption As String)
'���ܣ�ִ��SQL���
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_��Ԫ����.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Private Function ������ϸд��(ByVal rs��ϸ As ADODB.Recordset, Optional ByVal bln���� As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ���ϸ��¼
    '--�����:
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutPut As String
    Dim str��ϸ As String
    
    Dim strArr
    
    ������ϸд�� = False
    g�������_ǭ��.�����ܶ� = 0
    
    Err = 0:    On Error GoTo ErrHand:
    'Ȼ����봦����ϸ
    With rs��ϸ
        If .RecordCount = 0 Then
            ShowMsgbox "��������ص���ϸ���ü�¼!"
            Exit Function
        End If
        'YBJGBH  PCHAR   ���ջ������
        'CFH PCHAR   ������
        'CFMXDATA    PCHAR   ������ϸ����    ��ʽ˵��������1(ҽ��ҩƷ���+�м����+����+�м��������+)+�м����+
        'CPASSWORD   PCHAR   �ֿ��˿�����
        'CCZYXM  PCHAR   ����Ա����
        strInput = g�������_��Ԫ����.��������
        strInput = strInput & vbTab & Nvl(!no)
        
        Do While Not rs��ϸ.EOF
            gstrSQL = "Select * From ҽ��֧����Ŀ where ����=" & gintInsure & " and ����=" & g�������_��Ԫ����.�籣���� & " and �շ�ϸĿid=" & Nvl(!�շ�ϸĿID, 0)
            Call OpenRecordset(rsTemp, "ȷ��ҽ��֧����Ŀ")
            If rsTemp.EOF Then
                gstrSQL = "Select * From �շ�ϸĿ where id=" & Nvl(!�շ�ϸĿID, 0)
                If rsTemp.EOF Then
                    ShowMsgbox "��������ص��շ���Ŀ!"
                Else
                    ShowMsgbox "���շ���Ŀ�У���ĿΪ:" & rsTemp!���� & "δ������ض���!"
                End If
                Exit Function
            End If
            If Val(Nvl(rs��ϸ("ʵ�ս��"), 0)) <> 0 Then
                str��ϸ = str��ϸ & "@$" & Nvl(rsTemp!��Ŀ����)
                str��ϸ = str��ϸ & "||" & Nvl(rsTemp!��Ŀ����)
                str��ϸ = str��ϸ & "||" & Nvl(rsTemp!����, 0)
                str��ϸ = str��ϸ & "||" & Nvl(rsTemp!����, 0)
            End If
            g�������_��Ԫ����.�����ܶ� = g�������_��Ԫ����.�����ܶ� + Nvl(!ʵ�ս��, 0)
            
            'д�ϴ���־
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,null)"
            ExecuteProcedure "�����ϴ���־"
            
            rs��ϸ.MoveNext
        Loop
    End With
    str��ϸ = Mid(str��ϸ, 3)
    strInput = strInput & vbTab & str��ϸ
    strInput = strInput & vbTab & g�������_��Ԫ����.����
    strInput = strInput & vbTab & gstrUserName
    
    If ҵ������_��Ԫ����(�����ʻ�����_����, strInput, strOutPut) = False Then Exit Function
    If strOutPut = "" Then Exit Function
    strArr = Split(strOutPut, "||")
    
    With g��������
        .���� = strArr(0)
        .���� = strArr(1)
        .����ǰ�ʻ���� = Val(strArr(2))
        .�����ʻ�֧����� = Val(strArr(3))
        .�Էѽ�� = Val(strArr(4))
        .���Ѻ��ʻ���� = Val(strArr(5))
        .����ʱ�� = strArr(6)
        .ǰ�˵��ݺ� = strArr(7)
        .���ĵ��ݺ� = strArr(8)
        .������ = strArr(9)
        .����Ա���� = strArr(10)
        .ǰ������ = strArr(11)
    End With
    ������ϸд�� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function ���㷽ʽ����() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��������ʾ������
    '--�����:
    '--������:str���㷽ʽ
    '--��  ��:�ɹ�����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String
    Dim dbl�����ܶ� As Double
        
    '�����ܶ�=�����Էѽ��+����ͳ��֧�����+��ͳ����      �˽����������˺�������湫ʽת��������
    
    '�����Էѽ�� = �ܷ��ö� - ����ͳ��֧����� - �� / �߶�ͳ��֧�����
    '�Էѽ��ֽ�֧����ʻ�֧���� (��:��ѡ�����ֽ�����ʻ�֧��)
    '��ͳ����߶�ͳ��������ͬ
    'ͳ��֧��������ҽ���ڷ��ø��ݲ�ͬ���𸶱�׼�ͱ���������ҽ��������
    '��˵�����ݱ��������漼�������ɷ����޹�˾�������Ľ���
    ���㷽ʽ���� = False
    
    Err = 0:    On Error GoTo ErrHand:
    DebugTool "����(" & "Get���㷽ʽ" & ")"
    
    '����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
    dbl�����ܶ� = g��������.�����ʻ�֧����� + g��������.�Էѽ��
    str���㷽ʽ = "||�����ʻ�|" & g��������.�����ʻ�֧�����
    
    If Format(g�������_��Ԫ����.�����ܶ�, "###0.00;-###0.00;0;0") <> Format(dbl�����ܶ�, "###0.00;-###0.00;0;0") Then
        Dim blnYes As Boolean
        '�����ܶ���ҽ�����ķ����ܶ��,���ܽ��н���
        ShowMsgbox "���ν����ܶ�(" & g�������_��Ԫ����.�����ܶ� & ") ��" & vbCrLf & _
                    "   ���ķ��ص��ܶ�(" & dbl�����ܶ� & ")���²��ܽ���?"
        Exit Function
    End If
    
   '�������,�򱣴��Ԥ����¼��
    If str���㷽ʽ <> "" Then
        str���㷽ʽ = Mid(str���㷽ʽ, 3)
        g�������_�ɶ��ڽ�.���㷽ʽ = str���㷽ʽ
        
        If g��������.�����־ = 0 Then
            gstrSQL = "zl_���˽����¼_Update(" & g��������.����id & ",'" & str���㷽ʽ & "', 0)"
            Call ExecuteProcedure("����Ԥ����¼")
        Else
                gstrSQL = "zl_���˽����¼_Update(" & g��������.����id & ",'" & str���㷽ʽ & "',1)"
                Call ExecuteProcedure("����Ԥ����¼")
        End If
    End If
    
    '��ʾ������Ϣ
    If frm������Ϣ.ShowME(g��������.����id, False, "�����ʻ�:" & g��������.�����ʻ�֧�����, IIf(g��������.�����־ = 0, 0, 1)) = False Then
        ���㷽ʽ���� = False
        Exit Function
    End If
    ���㷽ʽ���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ��ȡ�����ʻ�֧��() As Double
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ�����ʻ�ֵ(��Ԥ����¼�л�ȡ)
    '--�����:
    '--������:
    '--��  ��:�ɹ�,���ر��θ����ʻ�֧��,���򷵻�0
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select ��� From ����Ԥ����¼ where ����ID=" & g��������.����id & " and  ���㷽ʽ='�����ʻ�'"
    
    OpenRecordset rsTemp, "��ȡ�����ʻ�֧��"
    If Not rsTemp.EOF Then
        ��ȡ�����ʻ�֧�� = Nvl(rsTemp!���, 0)
    End If
    
End Function
