Attribute VB_Name = "mdl�ɶ�����"
Option Explicit
Public Enum ҵ������_�ɶ�����
    ����籣���� = 0
    ��òα���Ա����
    ��Ժ�Ǽ�
    ȡ����Ժ�Ǽ�
    ��Ժ�Ǽ�
    ȡ����Ժ�Ǽ�
    ���Ӵ�������
    ���Ӵ�����ϸ
    ɾ���������ݼ�����ϸ
    ������������
    ��Ժ����
    ȡ����Ժ����
    
    ��ӡ��Ժ���㱨����
    ��ӡסԺ��Ա������㵥
    ������Ա��������
    ��ȡ��������
    ��ȡסԺ��¼��
    ���κ�����
    �����κ�����
    �Ͽ��κ�����
    ��ȡҩƷ��Ϣ
End Enum
Private Type InitbaseInfor
    ģ������ As Boolean                     '��ǰ�Ƿ���ģ���ȡҽ���ӿ�����
    ҽԺ���� As String                      '��ʼҽԺ����
    �������� As String                      'Ĭ�ϵ��籣��������
    
End Type
Public InitInfor_�ɶ����� As InitbaseInfor

Private Type �������
        ��¼��        As String
        ���Ϻ�    As String       '��ҽ����
        ����     As String
        �Ա�     As String
        ��������  As String
        ����        As Integer
        ҽ������    As String
        ���ݹ���    As String
        ��λ����    As String
        ��λ����    As String
        ҽ�Ʊ�־    As String
        ��������    As String
        
        �����ܶ�    As Double
        ����ID      As Long
        ���ֱ���    As String
        ��������    As String
End Type
Private Type ��������
    ҽ������ As Double
    �����㸶�� As Double
End Type
Private g������� As ��������
Public g�������_�ɶ����� As �������
Public gcnOracle_�ɶ����� As ADODB.Connection     '�м������
Private gbln������� As Boolean
Private gbln�Ѿ���ʼ As Boolean             '�Ѿ�����ʼ����.
'1.����籣������ź������б�
Private Declare Function GetSBJGLB Lib "cdgk_Yb.dll" () As String
'===============================================================================================================
'ԭ��:FUNCTION GETSBJGLB:PCHAR
'����: ����籣������ź������б�
'��ڲ���: ��
'���ڲ���: ��
'����: A�籣�������+�м����+A�籣��������+�м����+B�籣�������+�м����+B�籣��������+����
'===============================================================================================================

'2����òα���Ա�Ļ�������
Private Declare Function GetRYJBZL Lib "cdgk_Yb.dll" (ByVal str���Ϻ� As String, ByVal str�籣��� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GETRYJBZL(ASBBH,ABXJGBH:PCHAR):PCHAR;
'����: ��òα���Ա�Ļ�������
'��ڲ���: ASBBH   PCHAR   �α���Ա����ᱣ�Ϻ�
'          ABXJGBH PCHAR   �α���Ա���ڵı��ջ������
'���ڲ���: ��
'����: A�籣�������+�м����+A�籣��������+�м����+B�籣�������+�м����+B�籣��������+����
'===============================================================================================================

'3.��Ժ�Ǽ�
Private Declare Function RYDJ Lib "cdgk_Yb.dll" (ByVal strסԺ�� As String, ByVal str�������� As String, ByVal str������� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION RYDJ(AZYH,;ARYZL,ABXJGBH:PCHAR):PCHAR;
'����: ���籣����ҽ�����ݿ��ҽԺ����ҽ�����ݿ��ж�סԺ��ҽ�����˽��еǼǡ�
'��ڲ���: strסԺ��   PCHAR   סԺ��
'          str�������� PCHAR   �α���Ա�ĸ�������
'          str������� PCHAR �α���Ա���ڵ��籣�������
'���ڲ���: ��
'����:���ر�־@$��ᱣ�Ϻ�||���˼�¼��||ҽ������||���ݹ���||��λ����||����||�Ա�||�������ڣ���ʽ��YYYY-MM-DD��||��λ���||�μӻ���ҽ�Ʊ�־||��Ժ���ڣ���ʽ��YYYY-MM-DD��||���ֱ��||��������||����
'===============================================================================================================

'4.ȡ��סԺ
Private Declare Function ZYQX Lib "cdgk_Yb.dll" (ByVal strסԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION ZYQX(AZYH:PCHAR):PCHAR
'����: ���籣����ҽ�����ݿ��ҽԺ����ҽ�����ݿ���ɾ��ҽ������סԺ��¼��
'��ڲ���: strסԺ��   PCHAR   סԺ��
'���ڲ���: ��
'����:���ر�־
'===============================================================================================================

'5.��Ժ�Ǽ�
Private Declare Function CYCS Lib "cdgk_Yb.dll" (ByVal strסԺ�� As String, ByVal str��Ժ���� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION CYCS(AZYH ,CYRQ:PCHAR):PCHAR;
'����: ��ҽ������סԺ���������������ϴ����籣����ҽ�����ݿ⣻�Ա���ҽ�����ݿ���ҽ����������Ժ����
'��ڲ���: strסԺ��   PCHAR   סԺ��
'          str��Ժ���� pchar ��Ժ���ڣ�YYYY-MM-DD��
'���ڲ���: ��
'����:���ر�־
'===============================================================================================================

'6.ȡ����Ժ�Ǽ�
Private Declare Function CYCSQX Lib "cdgk_Yb.dll" (ByVal strסԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION CYCSQX (AZYH:PCHAR):PCHAR;
'����:ȡ���α��������籣���Ѿ�����ĳ�Ժ���ݣ��Ա����´��䡣
'��ڲ���: strסԺ��   PCHAR   סԺ��
'���ڲ���: ��
'����:'OK'�������Ϣ
'===============================================================================================================


'7.����һ����������
Private Declare Function AddCFJL Lib "cdgk_Yb.dll" (ByVal strסԺ�� As String, ByVal str�������� As String, ByVal strҽ�� As String, ByVal str���� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION ADDCFJL(AZYH,ACFRQ,AYS,AKS:PCHAR):PCHAR
'����:����һ���������ݡ���
'��ڲ���:
'        AZYH    PCHAR   סԺ��
'        ACFRQ   PCHAR   �������ڣ�YYYY-MM-DD��
'        AYS PCHAR   ҽ��
'        AKS PCHAR   ����
'���ڲ���: ��
'����:'OK'+�м����+������¼�Ż������Ϣ
'===============================================================================================================

'7.���Ӵ�����ϸ
Private Declare Function AddCFMX Lib "cdgk_Yb.dll" (ByVal str������¼�� As String, ByVal strҽ������ As String, ByVal str���� As String, ByVal str���� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION ADDCFMX(ACFID,AYPBH,ASL,ADJ:PCHAR):PCHAR;
'����:����һ��������ϸ��
'��ڲ���:
'    ACFID   PCHAR   ������¼��
'    AYPBH   PCHAR   ҩƷ���(�籣ҩƷ���)
'    ASL PCHAR   ����(����Ϊ����)
'    ADJ PCHAR   ����
'���ڲ���: ��
'����:'OK'+�м����+������ϸ��¼��+�м����+�Էѱ���+�м����+�Էѽ��������Ϣ
'===============================================================================================================

'8.ɾ���������ݼ�����ϸ
Private Declare Function DELCFJL Lib "cdgk_Yb.dll" (ByVal str������¼�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION DELCFJL(ACFID:PCHAR):PCHAR
'����:ɾ���������ݼ�����������ϸ��¼��
'��ڲ���:
'    ACFID   PCHAR   ������¼��
'���ڲ���: ��
'����:'OK'�������Ϣ
'===============================================================================================================


'9.������������
Private Declare Function CFCS Lib "cdgk_Yb.dll" (ByVal strסԺ�� As String, ByVal str������¼�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION CFCS(AZYH:PCHAR;ACFID:PCHAR):PCHAR
'����:���籣����ÿ��Ĵ���������籣���������ݿ⴫�䣨ͬһ���������Զ���ظ����䣬��һ�δ�������ݽ�����ǰһ�δ�������ݣ�
'��ڲ���:
'    AZYH    PCHAR   סԺ��
'    ACFID   PCHAR   ������¼�ţ�ͨ������ADDCFJL���ص�ֵ��
'���ڲ���: ��
'����:'OK'�������Ϣ
'===============================================================================================================

'10.��Ժ����
Private Declare Function CYJS Lib "cdgk_Yb.dll" (ByVal strסԺ�� As String, ByVal strԤ���־ As String) As String
'===============================================================================================================
'ԭ��:FNCTION CYJS(AZYH:PCHAR; ISPREV:INTEGER):PCHAR
'����:סԺ�α����˳�Ժ��סԺ��Ԥ����
'��ڲ���:
'    AZYH    PCHAR   סԺ��
'    ISPREV  PCHAR   Ԥ�����־��'0'����ʾԤ���㣩
'���ڲ���: ��
'����:'OK'�������Ϣ
'===============================================================================================================

'11.ȡ����Ժ����
Private Declare Function CYJSQX Lib "cdgk_Yb.dll" (ByVal strסԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION CYJSQX(AZYH:PCHAR):PCHAR
'����:ȡ���α����˳�Ժ����
'��ڲ���:
'    AZYH    PCHAR   סԺ��
'���ڲ���: ��
'����:'OK'�������Ϣ
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

'16.���κ������Ƿ����ӳɹ�
Private Declare Function CheckCon Lib "cdgk_Yb.dll" () As String
'===============================================================================================================
'ԭ��:FUNCTION CHECKCON:PCHAR;
'����:���κ������Ƿ����ӳɹ�
'��ڲ���:
'����:OK�������Ϣ
'===============================================================================================================

'17.�����κ�����
Private Declare Function RasDial Lib "cdgk_Yb.dll" (ByVal str�������� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION RASDIAL(SBXJGBH:PCHAR):PCHAR
'����:�κ���ѡ����籣�֣����佨������
'��ڲ���:SBXJGBH PCHAR   ���ջ������
'����:  �ɹ�    ������HIS�κ���״̬����ʾ"����"
'       ʧ�� ������Ϣ
'===============================================================================================================

'18.�Ͽ����籣�ֵ�����
Private Declare Function DisDial Lib "cdgk_Yb.dll" () As String
'===============================================================================================================
'ԭ��:FUNCTION DISDIAL:PCHAR
'����:�κ���ѡ����籣�֣����佨������
'��ڲ���:
'����:
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




Public Function ҽ����ʼ��_�ɶ�����() As Boolean
    Dim strReg As String, strOutPut As String
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    '��ʼģ��ӿ�
    Call GetRegInFor(g����ģ��, "����", "ģ��ӿ�", strReg)
    If Val(strReg) = 1 Then
        InitInfor_�ɶ�����.ģ������ = True
    Else
        InitInfor_�ɶ�����.ģ������ = False
    End If
    
   Call GetRegInFor(g����ģ��, "ҽ��", "�籣��������", strReg)
   
   InitInfor_�ɶ�����.�������� = strReg
   If strReg = "" Then
        MsgBox "��δ����Ĭ�ϵ��籣�������룬�����������!"
        Exit Function
   End If
   
    'ȡҽԺ����
    gstrSQL = "Select ҽԺ���� From ������� Where ���=" & TYPE_�ɶ�����
    Call OpenRecordset(rsTemp, "��ȡҽԺ����")
    InitInfor_�ɶ�����.ҽԺ���� = Nvl(rsTemp!ҽԺ����)
    
    
    '�м������
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=" & TYPE_�ɶ�����
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
    Set gcnOracle_�ɶ����� = New ADODB.Connection

    If OraDataOpen(gcnOracle_�ɶ�����, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ�ҽ���м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
    
   '�����κ�����
   If gbln�Ѿ���ʼ = False And gbln������� Then
        If ҵ������_�ɶ�����(�����κ�����, InitInfor_�ɶ�����.��������, strOutPut) = False Then
             Exit Function
        End If
   End If
   
   If gbln������� Then
        '���κ�����
        If ҵ������_�ɶ�����(���κ�����, "", strOutPut) = False Then
             Exit Function
        End If
    End If
    gbln�Ѿ���ʼ = True
    ҽ����ʼ��_�ɶ����� = True
End Function

Public Function ҽ����ֹ_�ɶ�����() As Boolean
    Dim strOutPut As String
    
    If gcnOracle_�ɶ�����.State = 1 Then
        gcnOracle_�ɶ�����.Close
    End If
    '�����κ�����
   Call ҵ������_�ɶ�����(�Ͽ��κ�����, "", strOutPut)
    Err = 0
    On Error Resume Next
    ҽ����ֹ_�ɶ����� = True
End Function

Public Function ��ݱ�ʶ_�ɶ�����(Optional bytType As Byte, Optional lng����ID As Long) As String
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    '���أ��ջ���Ϣ��
    Err = 0
    On Error GoTo ErrHand:
    If bytType = 0 Or bytType = 3 Then Exit Function
    
    ��ݱ�ʶ_�ɶ����� = frmIdentify�ɶ�����.GetPatient(bytType, lng����ID)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_�ɶ����� = ""
End Function


Public Function �������_�ɶ�����(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(�ʻ����,0) as �ʻ���� from �����ʻ� where ����ID='" & lng����ID & "' and ����=" & TYPE_�ɶ�����
    Call OpenRecordset(rsTemp, "��ȡ�����ʻ����")
    
    If rsTemp.EOF Then
        �������_�ɶ����� = 0
    Else
        �������_�ɶ����� = rsTemp("�ʻ����")
    End If
End Function
Public Function �����������_�ɶ�����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    �����������_�ɶ����� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_�ɶ�����(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
        '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    �������_�ɶ����� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ����������_�ɶ�����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    ����������_�ɶ����� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_�ɶ�����(lng����ID As Long, lng��ҳid As Long, ByRef strҽ���� As String) As Boolean
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    
    Dim rsTemp As New ADODB.Recordset, rsData As New ADODB.Recordset
    Dim strOutPut As String, strInPut As String
    Dim strArr
    Err = 0: On Error GoTo ErrHand:
    
    '��ȡסԺ��
    gstrSQL = "Select ҽ��סԺ��_ID.nextval  as סԺ��  From dual "
    OpenRecordset_�ɶ����� rsTemp, "��ȡסԺ��"
    
    
    
    'סԺ��||��������||�籣�������
    strInPut = Lpad(Nvl(rsTemp!סԺ��), 8)
    strInPut = strInPut & "||" & Get��������(lng����ID, lng��ҳid)
    strInPut = strInPut & "||" & g�������_�ɶ�����.��������
    If ҵ������_�ɶ�����(��Ժ�Ǽ�, strInPut, strOutPut) = False Then
        Exit Function
    End If
    
    '��ᱣ�Ϻ�||���˼�¼��||ҽ������||���ݹ���||��λ����||����||�Ա�||�������ڣ���ʽ��YYYY-MM-DD��||��λ���||�μӻ���ҽ�Ʊ�־||��Ժ���ڣ���ʽ��YYYY-MM-DD��||���ֱ��||��������||����
    strArr = Split(strOutPut, "||")
    '������ص���Ϣ
    ''OK'+�м����+�籣����סԺ��¼��
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & gintInsure & ",'ҽ��סԺ��','''" & Val(Nvl(rsTemp!סԺ��)) & "''')"
    Call ExecuteProcedure("ҽ��סԺ��")
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & gintInsure & ",'סԺ��¼��','''" & Val(strArr(0)) & "''')"
    Call ExecuteProcedure("����סԺ��¼��")
'    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & gintInsure & ",'��������','''" & strArr(12) & "''')"
'    Call ExecuteProcedure("��������")
    
    '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    ��Ժ�Ǽ�_�ɶ����� = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_�ɶ����� = False
End Function
Private Function Get��������(ByVal lng����ID As Long, ByVal lng��ҳid As Long) As String
    '    ��ᱣ�Ϻ�|���˼�¼��|ҽ������|���ݹ���|��λ����|����|�Ա�|�������ڣ���ʽ��YYYY-MM-DD��
    '    ��λ���|�μӻ���ҽ�Ʊ�־|��Ժ���ڣ���ʽ��YYYY-MM-DD��|���ֱ��|��������|����
    Dim rsTemp As New ADODB.Recordset
    Dim strInPut As String
    gstrSQL = "" & _
        "   Select  to_char(a.��Ժ����,'yyyy-mm-dd') as ��Ժ����,b.���� as ����" & _
        "   From ������ҳ a,���ű� b " & _
        "   Where A.��Ժ����ID=b.id(+) and a.����id=" & lng����ID & " and a.��ҳid=" & lng��ҳid
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��ҳ��Ϣ"
    With g�������_�ɶ�����
        strInPut = .���Ϻ�
        strInPut = strInPut & vbTab & "|" & .��¼��
        strInPut = strInPut & vbTab & "|" & .ҽ������
        strInPut = strInPut & vbTab & "|" & .���ݹ���
        strInPut = strInPut & vbTab & "|" & .��λ����
        strInPut = strInPut & vbTab & "|" & .����
        strInPut = strInPut & vbTab & "|" & .�Ա�
        strInPut = strInPut & vbTab & "|" & .��������
        strInPut = strInPut & vbTab & "|" & .��λ����
        strInPut = strInPut & vbTab & "|" & .ҽ�Ʊ�־
        strInPut = strInPut & vbTab & "|" & Nvl(rsTemp!��Ժ����)
        strInPut = strInPut & vbTab & "|" & .���ֱ���
        strInPut = strInPut & vbTab & "|" & .��������
        strInPut = strInPut & vbTab & "|" & Nvl(rsTemp!����)
    End With
    Get�������� = strInPut
    
    
End Function
Private Function Get���״���(ByVal intType As ҵ������_�ɶ�����, Optional bln������ As Boolean = False) As String
    '������û��
    Select Case intType
        Case ����籣����
            Get���״��� = IIf(bln������, "����籣����", "01")
        Case ��òα���Ա����
            Get���״��� = IIf(bln������, "��òα���Ա����", "02")
        Case ��Ժ�Ǽ�
            Get���״��� = IIf(bln������, "��Ժ�Ǽ�", "03")
        Case ȡ����Ժ�Ǽ�
            Get���״��� = IIf(bln������, "ȡ����Ժ�Ǽ�", "04")
        Case ��Ժ�Ǽ�
            Get���״��� = IIf(bln������, "��Ժ�Ǽ�", "05")
        Case ȡ����Ժ�Ǽ�
            Get���״��� = IIf(bln������, "ȡ����Ժ�Ǽ�", "06")
        Case ���Ӵ�������
            Get���״��� = IIf(bln������, "���Ӵ�������", "07")
        Case ���Ӵ�����ϸ
            Get���״��� = IIf(bln������, "���Ӵ�����ϸ", "08")
        Case ɾ���������ݼ�����ϸ
            Get���״��� = IIf(bln������, "ɾ���������ݼ�����ϸ", "09")
        Case ������������
            Get���״��� = IIf(bln������, "������������", "10")
        Case ��Ժ����
            Get���״��� = IIf(bln������, "��Ժ����", "11")
        Case ȡ����Ժ����
            Get���״��� = IIf(bln������, "ȡ����Ժ����", "12")
        Case ��ӡ��Ժ���㱨����
            Get���״��� = IIf(bln������, "��ӡ��Ժ���㱨����", "13")
        Case ��ӡסԺ��Ա������㵥
            Get���״��� = IIf(bln������, "��ӡסԺ��Ա������㵥", "14")
        Case ������Ա��������
            Get���״��� = IIf(bln������, "������Ա��������", "15")
        Case ��ȡ��������
            Get���״��� = IIf(bln������, "��ȡ��������", "16")
        Case ��ȡסԺ��¼��
            Get���״��� = IIf(bln������, "��ȡסԺ��¼��", "17")
        Case ���κ�����
            Get���״��� = IIf(bln������, "���κ�����", "18")
        Case �����κ�����
            Get���״��� = IIf(bln������, "�����κ�����", "19")
        Case �Ͽ��κ�����
            Get���״��� = IIf(bln������, "�Ͽ��κ�����", "20")
        Case ��ȡҩƷ��Ϣ
            Get���״��� = IIf(bln������, "��ȡҩƷ��Ϣ", "21")
        Case Else
            Get���״��� = IIf(bln������, "����Ľ��״���", "-1")
    End Select
End Function
Public Function ҵ������_�ɶ�����(ByVal intType As ҵ������_�ɶ�����, strInputString As String, strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ҵ�����ҵ������
    '--�����:strinPutString-���봮,������˳��,��tab���ָ��Ĵ��봮
    '--������:strOutPutString-�����,������˳��,��tab���ָ��ķ��ش�
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strInPut As String, lngReturn As Long, strOutPut As String, strReturn As String
    Dim strInValue(0 To 20) As String
    
    Dim str���״��� As String
    Dim i As Integer
    Dim strArr
    
    str���״��� = Get���״���(intType)
    strInPut = str���״��� & "|" & strInputString
    DebugTool "����ҵ��������(ҵ������Ϊ:" & intType & "),�������Ϊ" & vbCrLf & str���״��� & "|" & strInPut
    
    ҵ������_�ɶ����� = False
    If InitInfor_�ɶ�����.ģ������ Then
        '��ȡģ������
        Readģ������ intType, strInPut, strOutPutstring
         ҵ������_�ɶ����� = True
        Exit Function
    End If
    strArr = Split(strInputString, "||")
    For i = 0 To UBound(strArr)
        strInValue(i) = strArr(i)
    Next
    
    Err = 0
    On Error GoTo ErrHand:
    
    Select Case intType
        Case ����籣����
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
        Case ��òα���Ա����
            strOutPut = GetRYJBZL(strInValue(0), strInValue(1))
            If strOutPut = "" Then
                MsgBox "��òα���Ա����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = strArr(1)
        Case ��Ժ�Ǽ�
            '
            strOutPut = RYDJ(strInValue(0), Replace(strInValue(1), vbTab & "|", "||"), strInValue(2))
            If strOutPut = "" Then
                MsgBox "��Ժ�Ǽ�ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = strArr(1)
        Case ȡ����Ժ�Ǽ�
            strOutPut = ZYQX(strInValue(0))
            If strOutPut = "" Then
                MsgBox "ȡ����Ժ�Ǽ�ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = ""
        Case ��Ժ�Ǽ�
            strOutPut = CYCS(strInValue(0), strInValue(1))
            If strOutPut = "" Then
                MsgBox "��Ժ�Ǽ�ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = ""
        Case ȡ����Ժ�Ǽ�
            strOutPut = CYCSQX(strInValue(0))
            If strOutPut = "" Then
                MsgBox "ȡ����Ժ�Ǽ�ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = ""
        Case ���Ӵ�������
            strOutPut = AddCFJL(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            If strOutPut = "" Then
                MsgBox "���Ӵ�������ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = strArr(1)
        Case ���Ӵ�����ϸ
            strOutPut = AddCFMX(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            If strOutPut = "" Then
                MsgBox "���Ӵ�����ϸʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = ""
            For i = 1 To UBound(strArr)
                strOutPut = "||" & strArr(i)
            Next
            If strOutPut <> "" Then
                strOutPut = Mid(strOutPut, 3)
            End If
        Case ɾ���������ݼ�����ϸ
            strOutPut = DELCFJL(strInValue(0))
            If strOutPut = "" Then
                MsgBox "ɾ���������ݼ�����ϸʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = ""
        Case ������������
            strOutPut = CFCS(strInValue(0), strInValue(1))
            If strOutPut = "" Then
                MsgBox "������������ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = ""
        Case ��Ժ����
            strOutPut = CFCS(strInValue(0), strInValue(1))
            If strOutPut = "" Then
                MsgBox "��Ժ����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = strArr(1)
        Case ȡ����Ժ����
            strOutPut = CYJSQX(strInValue(0))
            If strOutPut = "" Then
                MsgBox "��Ժ����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = ""
        Case ��ӡ��Ժ���㱨����
            strOutPut = JSReport(strInValue(0), strInValue(1))
            strOutPut = ""
        Case ��ӡסԺ��Ա������㵥
            strOutPut = CWJSReport(strInValue(0), strInValue(1))
            strOutPut = ""
        
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
        Case ���κ�����
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
        Case �����κ�����
            strOutPut = RasDial(strInValue(0))
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = ""
        Case �Ͽ��κ�����
            strOutPut = DisDial()
            strOutPut = ""
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
    ҵ������_�ɶ����� = True
    DebugTool "ҵ������ɹ�(ҵ������Ϊ:" & intType & ")." & vbCrLf & "�������Ϊ" & vbCrLf & strInputString & vbCrLf & "�������Ϊ:" & vbCrLf & strReturn
     Exit Function
    
ErrHand:
    DebugTool "ҵ������ʧ��(ҵ������Ϊ:" & intType & ")." & vbCrLf & "�������Ϊ" & vbCrLf & strInputString & vbCrLf & "�������Ϊ:" & vbCrLf & strReturn
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_�ɶ�����(lng����ID As Long, lng��ҳid As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
            
    '���˺�:20040923���ӵ�
    Dim rsTemp As New ADODB.Recordset
    Dim strInPut As String, strOutPut As String
    
    Err = 0
    On Error GoTo ErrHand
    
    DebugTool "������Ժ�ǳ����ӿ�"
    
    ��Ժ�Ǽǳ���_�ɶ����� = False
    If ����δ�����(lng����ID, lng��ҳid) Then
        ShowMsgbox "����δ����ã����ܳ�����Ժ�Ǽ�"
        Exit Function
    End If
    
    '��ȡסԺ��
    gstrSQL = "Select ҽ��סԺ�� סԺ�� From �����ʻ� where ����ID=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��Ժ�Ǽǳ���"
    If ҵ������_�ɶ�����(ȡ����Ժ�Ǽ�, Lpad(Nvl(rsTemp!סԺ��), 8), strOutPut) = False Then Exit Function

    '����ҽ���ʻ�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ɶ����� & ")"
    Call ExecuteProcedure("��������Ժ�Ǽ�")
    DebugTool "ȡ���ɹ�"
    ��Ժ�Ǽǳ���_�ɶ����� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_�ɶ�����(lng����ID As Long, lng��ҳid As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset
    Dim strInPut As String, strOutPut As String
    
    Err = 0
    On Error GoTo ErrHand:
    
    gstrSQL = "" & _
        "   Select B.ҽ��סԺ�� סԺ��,to_Char(a.��Ժ����,'yyyy-MM-DD') as ��Ժ����" & _
        "   From ������ҳ A,�����ʻ� B " & _
        "   Where A.����iD=b.����id " & _
        "       and A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳid
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡסԺ�źͳ�Ժ����"
    If rsTemp.EOF Then
        ShowMsgbox "�޶�Ӧ��סԺ��Ա��Ϣ"
        Exit Function
    End If
        
    strInPut = Lpad(Nvl(rsTemp!סԺ��), 8)
    strInPut = strInPut & "||" & Nvl(rsTemp!��Ժ����)
    If ҵ������_�ɶ�����(��Ժ�Ǽ�, strInPut, strOutPut) = False Then Exit Function
        
    '�ı䵱ǰ״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    ��Ժ�Ǽ�_�ɶ����� = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_�ɶ����� = False
End Function
Public Function ��Ժ�Ǽǳ���_�ɶ�����(ByVal lng����ID As Long, ByVal lng��ҳid As Long) As Boolean
    '��Ժ�Ǽǳ���
    Dim rsTemp As New ADODB.Recordset
    Dim strInPut As String, strOutPut As String
    Dim strArr As Variant
    
     '�ı䲡��״̬
     If Not ����δ�����(lng����ID, lng��ҳid) Then
            ShowMsgbox "�ò����Ѿ���Ժ������,������ȡ����Ժ!"
            Exit Function
     End If
     
     gstrSQL = "Select ҽ��סԺ�� סԺ�� From �����ʻ� where ����ID=" & lng����ID
     zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡסԺ��"
     strInPut = Nvl(rsTemp!סԺ��)
     If ҵ������_�ɶ�����(ȡ����Ժ�Ǽ�, strInPut, strOutPut) = False Then Exit Function
     
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ɶ����� & ")"
    Call ExecuteProcedure("������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_�ɶ����� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_�ɶ�����(lng����ID As Long, ByVal lng����ID As Long) As Boolean
  '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    
    Dim rsTemp As New ADODB.Recordset, strInPut As String, strOutPut As String
    
    Dim lng��ҳid As Long
    Dim dbl�����ܶ� As Double
    Dim strArr
    Dim str���㷽ʽ  As String, strסԺ�� As String
    Dim obj���� As ��������
        
    Err = 0: On Error GoTo ErrHand:
    Call DebugTool("����סԺ����")
    
    
    If g�������_�ɶ�����.����ID <> lng����ID Then
        MsgBox "�ò���û�����ҽ����Ԥ������������ܽ��н��㡣", vbInformation, gstrSysName
        Exit Function
    End If
        
    gstrSQL = "Select ��ǰ״̬,ҽ��סԺ�� סԺ�� From �����ʻ� where ����ID=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "�жϵ�ǰ��סԺ״̬!"
    If Nvl(rsTemp!��ǰ״̬, 0) = 1 Then
        ShowMsgbox "��ǰ���˻�������Ժ״̬,���Ժ���ٽ���!"
        Exit Function
    End If
    strסԺ�� = Lpad(Nvl(rsTemp!סԺ��), 8)
    With g��������
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=" & lng����ID
        Call OpenRecordset(rsTemp, "�������")
        If IsNull(rsTemp("��ҳID")) = True Then
            MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
            Exit Function
        End If
        lng��ҳid = rsTemp("��ҳID")
    End With
    
    gstrSQL = " " & _
          " Select sum(nvl(���ʽ��,0)) as ʵ�ս�� " & _
          " From ���˷��ü�¼ " & _
          " Where ��¼״̬<>0 and ����ID=" & lng����ID & " and  Nvl(���ӱ�־,0)<>9"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�ܷ���"
    dbl�����ܶ� = Nvl(rsTemp!ʵ�ս��, 0)
    
    
    'AZYH    PCHAR   סԺ��
    'ISPREV  PCHAR   Ԥ�����־��'0'����ʾԤ���㣩
    strInPut = strסԺ��
    strInPut = strInPut & "||0"
    If ҵ������_�ɶ�����(��Ժ����, strInPut, strOutPut) = False Then Exit Function
    strArr = Split(strOutPut, "||")
    
    '����ֵ
    'Ӧ֧��ͳ���||���ⶥ�Ը�||�����Ը�С��||����֧���ϼ�||ͳ��֧��ͳ���||����Ӧ��֧��||���β����㸶��||���β����������||ʵ�ʿۼ�����||ͳ��ⶥ���||
    'ͳ���𸶽��||סԺ��¼��||���˼�¼��||����||סԺ��||���ֱ��||��������||����||ҽ�ƻ�����||��Ժ����||��Ժ����||�ѽ���ͳ���||����ҽ�Ʒ�С��||����ҩƷ��||
    '��������||�������Ʒ�||����������||�Ը�С��||�Ը�ҩƷ��||�Ը�����||�Ը����Ʒ�||�Ը�������||����ҩƷ��||��������||�������Ʒ�||����������||ͳ��֧������||
    'ͳ�����Ʒ�||ͳ��������||��Ժ��־||�����־||�����־||����ҽ��״̬||���㷽ʽ||��˷�ʽ||��������||��λ���||��λ����||��ᱣ�Ϻ�||����||�Ա�||��������||Ԥ�ɽ��||
    '����Ӧ�����||����ʵ�����||�˿���||����ʵ��֧�����||�籣������||�����������||��������־||���ջ�����||����Ա���||������ȡʱ��||ҽ�Ʊ��ձ��||�籣��������||
    '������������||�������ܱ�־||�����𸶿ۼ���־||����������||����������||�����㸶����||�������㸶��||�����㸶�����||����������ܿ�ʼ����||�㸶�ܶ�||
    With obj����
        .ҽ������ = Val(strArr(4))
        .�����㸶�� = Val(strArr(6))
    End With
    
    '�����������Ƿ�һ��
    With g�������
        If .ҽ������ <> obj����.ҽ������ Or .�����㸶�� <> obj����.�����㸶�� Then
            ShowMsgbox "���ν���ʱ��������㲻��,���������д������ϴ���,����..." & vbCrLf & _
                    "   ͳ��֧��:" & Format(.ҽ������, "####0.00;####0.00;0.00;0.00") & vbTab & vbTab & Format(obj����.ҽ������, "####0.00;####0.00;0.00;0.00") & vbCrLf & _
                    "   ͳ��֧��:" & Format(.�����㸶��, "####0.00;####0.00;0.00;0.00") & vbTab & vbTab & Format(obj����.�����㸶��, "####0.00;####0.00;0.00;0.00") & vbCrLf
            Exit Function
        End If
    End With
    
    '�ٴν���
  
    'AZYH    PCHAR   סԺ��
    'ISPREV  PCHAR   Ԥ�����־��'0'����ʾԤ���㣩
    strInPut = strסԺ��
    strInPut = strInPut & "||1"
    If ҵ������_�ɶ�����(��Ժ����, strInPut, strOutPut) = False Then Exit Function
    strArr = Split(strOutPut, "||")
    
    '����ֵ
    'Ӧ֧��ͳ���||���ⶥ�Ը�||�����Ը�С��||����֧���ϼ�||ͳ��֧��ͳ���||����Ӧ��֧��||���β����㸶��||���β����������||ʵ�ʿۼ�����||ͳ��ⶥ���||
    'ͳ���𸶽��||סԺ��¼��||���˼�¼��||����||סԺ��||���ֱ��||��������||����||ҽ�ƻ�����||��Ժ����||��Ժ����||�ѽ���ͳ���||����ҽ�Ʒ�С��||����ҩƷ��||
    '��������||�������Ʒ�||����������||�Ը�С��||�Ը�ҩƷ��||�Ը�����||�Ը����Ʒ�||�Ը�������||����ҩƷ��||��������||�������Ʒ�||����������||ͳ��֧������||
    'ͳ�����Ʒ�||ͳ��������||��Ժ��־||�����־||�����־||����ҽ��״̬||���㷽ʽ||��˷�ʽ||��������||��λ���||��λ����||��ᱣ�Ϻ�||����||�Ա�||��������||Ԥ�ɽ��||
    '����Ӧ�����||����ʵ�����||�˿���||����ʵ��֧�����||�籣������||�����������||��������־||���ջ�����||����Ա���||������ȡʱ��||ҽ�Ʊ��ձ��||�籣��������||
    '������������||�������ܱ�־||�����𸶿ۼ���־||����������||����������||�����㸶����||�������㸶��||�����㸶�����||����������ܿ�ʼ����||�㸶�ܶ�||
    With obj����
        .ҽ������ = Val(strArr(4))
        .�����㸶�� = Val(strArr(6))
    End With
    
    If InsertIntoҽ�������¼(strArr, lng����ID) = False Then Exit Function
    
    
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
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(),�ۼ�ͳ�ﱨ��_IN(),סԺ����_IN(��ҳID),����(��),�ⶥ��_IN(��),ʵ������_IN(��),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(��),�����Ը����_IN(��),
    '   ����ͳ����_IN(ͳ��֧��),ͳ�ﱨ�����_IN(ͳ��֧��),���Ը����_IN(�����㸶��),�����Ը����_IN(��),�����ʻ�֧��_IN(),"
    '   ֧��˳���_IN(סԺ��),��ҳID_IN(��ҳID),��;����_IN,��ע_IN()
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
   
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL,NULL,NULL," & lng��ҳid & ",0,0,0," & _
            dbl�����ܶ� & ",0,0," & _
            obj����.ҽ������ & "," & obj����.ҽ������ & ",0,0," & obj����.�����㸶�� & ",'" & _
            strסԺ�� & "'," & lng��ҳid & ",NULL,NULL)"
    Call ExecuteProcedure("��������¼")
    '---------------------------------------------------------------------------------------------
      
    סԺ����_�ɶ����� = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function InsertIntoҽ�������¼(ByVal strArr As Variant, ByVal lng����ID As Long) As Boolean
    '����:���м�����ҽ�������¼
    '����:strarr��split(stroutput,"||")����������
    
    Err = 0
    On Error GoTo ErrHand:
    InsertIntoҽ�������¼ = False
    
    DebugTool "����InsertIntoҽ�������¼"
    'strArr:
    'Ӧ֧��ͳ���||���ⶥ�Ը�||�����Ը�С��||����֧���ϼ�||ͳ��֧��ͳ���||����Ӧ��֧��||���β����㸶��||���β����������||ʵ�ʿۼ�����||ͳ��ⶥ���||
    'ͳ���𸶽��||סԺ��¼��||���˼�¼��||����||סԺ��||���ֱ��||��������||����||ҽ�ƻ�����||��Ժ����||��Ժ����||�ѽ���ͳ���||����ҽ�Ʒ�С��||����ҩƷ��||
    '��������||�������Ʒ�||����������||�Ը�С��||�Ը�ҩƷ��||�Ը�����||�Ը����Ʒ�||�Ը�������||����ҩƷ��||��������||�������Ʒ�||����������||ͳ��֧������||
    'ͳ�����Ʒ�||ͳ��������||��Ժ��־||�����־||�����־||����ҽ��״̬||���㷽ʽ||��˷�ʽ||��������||��λ���||��λ����||��ᱣ�Ϻ�||����||�Ա�||��������||Ԥ�ɽ��||
    '����Ӧ�����||����ʵ�����||�˿���||����ʵ��֧�����||�籣������||�����������||��������־||���ջ�����||����Ա���||������ȡʱ��||ҽ�Ʊ��ձ��||�籣��������||
    '������������||�������ܱ�־||�����𸶿ۼ���־||����������||����������||�����㸶����||�������㸶��||�����㸶�����||����������ܿ�ʼ����||�㸶�ܶ�||
    
    '���̲���
    '����,����ID,
    'Ӧ֧��ͳ���,���ⶥ�Ը�,�����Ը�С��,����֧���ϼ�,ͳ��֧��ͳ���,����Ӧ��֧��,���β����㸶��,���β����������,ʵ�ʿۼ�����,ͳ��ⶥ���,ͳ���𸶽��,סԺ��¼��,���˼�¼��,
    '����,סԺ��,���ֱ��,��������,����,ҽ�ƻ�����,��Ժ����,��Ժ����,�ѽ���ͳ���,����ҽ�Ʒ�С��,����ҩƷ��,��������,�������Ʒ�,����������,�Ը�С��,�Ը�ҩƷ��,�Ը�����,�Ը����Ʒ�,
    '�Ը�������,����ҩƷ��,��������,�������Ʒ�,����������,ͳ��֧������,ͳ�����Ʒ�,ͳ��������,��Ժ��־,�����־,�����־,����ҽ��״̬,���㷽ʽ,��˷�ʽ,��������,��λ���,��λ����,
    '��ᱣ�Ϻ�,����,�Ա�,��������,Ԥ�ɽ��,����Ӧ�����,����ʵ�����,�˿���,����ʵ��֧�����,�籣������,�����������,��������־,���ջ�����,����Ա���,������ȡʱ��,ҽ�Ʊ��ձ��,
    '�籣��������,������������,�������ܱ�־,�����𸶿ۼ���־,����������,����������,�����㸶����,�������㸶��,�����㸶�����,����������ܿ�ʼ����,�㸶�ܶ�
    
    '    ����        number(2),
    gstrSQL = "ZL_ҽ�������¼_INSERT(2"
    '    ����ID      number(18),
    gstrSQL = gstrSQL & "," & lng����ID
    '    Ӧ֧��ͳ���    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(0))
    '    ���ⶥ�Ը�  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(1))
    '    �����Ը�С��    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(2))
    '    ����֧���ϼ�    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(3))
    '    ͳ��֧��ͳ���  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(4))
    '    ����Ӧ��֧��    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(5))
    '    ���β����㸶��  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(6))
    '    ���β����������    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(7))
    '    ʵ�ʿۼ�����    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(8))
    '    ͳ��ⶥ���    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(9))
    '    ͳ���𸶽��    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(10))
    
    '    סԺ��¼��  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(11) & "'"
    '    ���˼�¼��  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(12) & "'"
    '    ����        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(13) & "'"
    '    סԺ��      varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(14) & "'"
    '    ���ֱ��        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(15) & "'"
    '    ��������        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(16) & "'"
    '    ����        varchar2(50),
    gstrSQL = gstrSQL & ",'" & strArr(17) & "'"
    '    ҽ�ƻ�����  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(18) & "'"
    '    ��Ժ����        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(19) & "'"
    '    ��Ժ����        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(20) & "'"
      
    '    �ѽ���ͳ���    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(21))
    '    ����ҽ�Ʒ�С��  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(22))
    '    ����ҩƷ��  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(23))
    '    ��������  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(24))
    '    �������Ʒ�  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(25))
    '    ����������  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(26))
    '    �Ը�С��        number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(27))
    '    �Ը�ҩƷ��  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(28))
    '    �Ը�����  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(29))
    '    �Ը����Ʒ�  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(30))
    '    �Ը�������  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(31))
    '    ����ҩƷ��  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(32))
    '    ��������  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(33))
    '    �������Ʒ�  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(34))
    '    ����������  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(35))
    '    ͳ��֧������    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(36))
    '    ͳ�����Ʒ�  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(37))
    '    ͳ��������  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(38))
      
    '    ��Ժ��־        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(39) & "'"
    '    �����־        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(40) & "'"
    '    �����־        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(41) & "'"
    '    ����ҽ��״̬    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(42) & "'"
    '    ���㷽ʽ        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(43) & "'"
    '    ��˷�ʽ        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(44) & "'"
    '    ��������        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(45) & "'"
    '    ��λ���        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(46) & "'"
    '    ��λ����        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(47) & "'"
    '    ��ᱣ�Ϻ�  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(48) & "'"
    '    ����        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(49) & "'"
    '    �Ա�        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(50) & "'"
    '    ��������        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(51) & "'"
        
    '    Ԥ�ɽ��        number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(52))
    '    ����Ӧ�����    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(53))
    '    ����ʵ�����    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(54))
    '    �˿���        number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(55))
    '    ����ʵ��֧�����    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(56))
    '    �籣������    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(57))
            
    '    �����������    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(58) & "'"
    '    ��������־    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(59) & "'"
    '    ���ջ�����  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(60) & "'"
    '    ����Ա���  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(61) & "'"
    '    ������ȡʱ��    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(62) & "'"
    '    ҽ�Ʊ��ձ��    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(63) & "'"
    '    �籣��������    varchar2(50),
    gstrSQL = gstrSQL & ",'" & strArr(64) & "'"
    '    ������������    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(65) & "'"
    '    �������ܱ�־    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(66) & "'"
    '    �����𸶿ۼ���־    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(67) & "'"
            
    '    ����������    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(68))
    '    ����������    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(69))
    '    �����㸶����    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(70))
    '    �������㸶��    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(71))
    '    �����㸶�����    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(72))
            
    '    ����������ܿ�ʼ����    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(73) & "'"
    '    �㸶�ܶ�        number(16,5))
    gstrSQL = gstrSQL & "," & Val(strArr(74)) & ")"
    gcnOracle_�ɶ�����.Execute gstrSQL, , adCmdStoredProc
    InsertIntoҽ�������¼ = True
    DebugTool "����ҽ�������¼�ɹ�"
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function סԺ�������_�ɶ�����(lng����ID As Long) As Boolean
     '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    Dim rs�����¼ As New ADODB.Recordset
    
    Dim strInPut As String, strOutPut  As String
    Dim lng����ID As Long, strסԺ�� As String
    Dim strArr
    
    Err = 0: On Error GoTo ErrHand:
    
    
    '�˷�
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=" & lng����ID
    Call OpenRecordset(rsTemp, "�������")
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=" & gintInsure & " and ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, "�������")
    If rsTemp.EOF = True Then
        MsgBox "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "select * from ҽ�������¼ where ����=2  and ����ID=" & lng����ID
    Call OpenRecordset_�ɶ�����(rs�����¼, "�������")
    If rsTemp.EOF = True Then
        MsgBox "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
        
    '�жϲ��˵�סԺ���������Ƿ��������ϡ��жϱ�׼�Ǽ�鲡�����µ�סԺ��¼������У��Ͳ��ܽ�����
    If CanסԺ�������(rsTemp("����ID"), rsTemp("��ҳID")) = False Then Exit Function
    
    strסԺ�� = rsTemp("֧��˳���")
    strInPut = strסԺ��
    If ҵ������_�ɶ�����(ȡ����Ժ����, strInPut, strOutPut) = False Then
        Exit Function
    End If
    
    'Ӧ֧��ͳ���||���ⶥ�Ը�||�����Ը�С��||����֧���ϼ�||ͳ��֧��ͳ���||����Ӧ��֧��||���β����㸶��||���β����������||ʵ�ʿۼ�����||ͳ��ⶥ���||ͳ���𸶽��||סԺ��¼��||���˼�¼��||����||סԺ��||���ֱ��||��������||����||ҽ�ƻ�����||��Ժ����||��Ժ����||�ѽ���ͳ���||����ҽ�Ʒ�С��||����ҩƷ��||��������||�������Ʒ�||����������||�Ը�С��||�Ը�ҩƷ��||�Ը�����||�Ը����Ʒ�||�Ը�������||����ҩƷ��||��������||�������Ʒ�||����������||ͳ��֧������||ͳ�����Ʒ�||ͳ��������||��Ժ��־||�����־||�����־||����ҽ��״̬||���㷽ʽ||��˷�ʽ||��������||��λ���||��λ����||��ᱣ�Ϻ�||����||�Ա�||��������||Ԥ�ɽ��||����Ӧ�����||����ʵ�����||�˿���||����ʵ��֧�����||�籣������||�����������||��������־||���ջ�����||����Ա���||������ȡʱ��||ҽ�Ʊ��ձ��||�籣��������||������������||�������ܱ�־||�����𸶿ۼ���־||����������||����������||�����㸶����||�������㸶��||�����㸶�����||����������ܿ�ʼ����||�㸶�ܶ�
    strArr = Split("Ӧ֧��ͳ���||���ⶥ�Ը�||�����Ը�С��||����֧���ϼ�||ͳ��֧��ͳ���||����Ӧ��֧��||���β����㸶��||���β����������||ʵ�ʿۼ�����||ͳ��ⶥ���||ͳ���𸶽��||סԺ��¼��||���˼�¼��||����||סԺ��||���ֱ��||��������||����||ҽ�ƻ�����||��Ժ����||��Ժ����||�ѽ���ͳ���||����ҽ�Ʒ�С��||����ҩƷ��||��������||�������Ʒ�||����������||�Ը�С��||�Ը�ҩƷ��||�Ը�����||�Ը����Ʒ�||�Ը�������||����ҩƷ��||��������||�������Ʒ�||����������||ͳ��֧������||ͳ�����Ʒ�||ͳ��������||��Ժ��־||�����־||�����־||����ҽ��״̬||���㷽ʽ||��˷�ʽ||��������||��λ���||��λ����||��ᱣ�Ϻ�||����||�Ա�||��������||Ԥ�ɽ��||����Ӧ�����||����ʵ�����||�˿���||����ʵ��֧�����||�籣������||�����������||��������־||���ջ�����||����Ա���||������ȡʱ��||ҽ�Ʊ��ձ��||�籣��������||������������||�������ܱ�־||�����𸶿ۼ���־||����������||����������||�����㸶����||�������㸶��||�����㸶�����||����������ܿ�ʼ����||�㸶�ܶ�", "||")
    
    strInPut = ""
    Dim i As Integer
    For i = 0 To UBound(strArr)
        If rs�����¼.Fields(strArr(i)).Type = 131 Then
            strInPut = strInPut & "||" & (Val(Nvl(rs�����¼.Fields(strArr(i)))) * -1)
        Else
            strInPut = strInPut & "||" & Nvl(rs�����¼.Fields(strArr(i)))
        End If
    Next
    strInPut = Mid(strInPut, 3)
    strArr = Split(strInPut, "||")
    If InsertIntoҽ�������¼(strArr, lng����ID) = False Then Exit Function
    
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(),�ۼ�ͳ�ﱨ��_IN(),סԺ����_IN(��ҳID),����(��),�ⶥ��_IN(��),ʵ������_IN(��),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(��),�����Ը����_IN(��),
    '   ����ͳ����_IN(ͳ��֧��),ͳ�ﱨ�����_IN(ͳ��֧��),���Ը����_IN(�����㸶��),�����Ը����_IN(��),�����ʻ�֧��_IN(),"
    '   ֧��˳���_IN(סԺ��),��ҳID_IN(��ҳID),��;����_IN,��ע_IN()
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & rsTemp("����ID") & "," & Year(zlDatabase.Currentdate) & "," & _
        "NULL,NULL,NULL,NULL," & Nvl(rsTemp!��ҳID, 0) & ",0,0,0," & _
        rsTemp("�������ý��") * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & "," & Nvl(rsTemp!���Ը����, 0) * -1 & ",0," & _
        "NULL,'" & strסԺ�� & "'," & Nvl(rsTemp!��ҳID, 0) & ",NULL,NULL)"
    Call ExecuteProcedure("����ҽ�������¼")
    
    סԺ�������_�ɶ����� = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function �����Ǽ�_�ɶ�����(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ�������ϸ����
    '--�����:
    '--������:
    '--��  ��:�ϴ��ɹ�����True,����False
    '-----------------------------------------------------------------------------------------------------------

    Dim lng����ID As Long
    Dim lng��ҳid As Long
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim strInPut As String, strOutPut As String
    Dim str������¼�� As String
    Dim strArr
    
    Err = 0
    On Error GoTo ErrHand:
    
    �����Ǽ�_�ɶ����� = False
    
   '�������ŵ��ݵķ�����ϸ
    gstrSQL = "Select A.ID,A.NO,M.ҽ��סԺ�� סԺ��,A.����ID,A.��ҳID,to_char(A.����ʱ��,'yyyy-mm-dd') as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս�� " & _
              "         ,A.�շ�ϸĿID,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
              "         ,Q.���� as ��������,C.��Ŀ����,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,A.����Ա����,B.���㵥λ " & _
              "  From ���˷��ü�¼ A,�շ�ϸĿ B,(select * From ����֧����Ŀ where ����=" & gintInsure & ") C,������ҳ D,�����ʻ� M,���ű� Q" & _
              "  where A.NO='" & str���ݺ� & "' and A.��¼����=" & lng��¼���� & " and A.����id=M.����id and a.��������ID=Q.id(+) and A.��¼״̬ = " & lng��¼״̬ & " And Nvl(A.�Ƿ��ϴ�,0)=0 " & _
              "        and A.����ID=D.����ID and A.��ҳID=D.��ҳID And D.����=" & gintInsure & _
              "        and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID(+) and D.����=" & gintInsure & _
              "  Order by A.����ID,A.NO,A.����ʱ��"
    
    Call OpenRecordset(rs��ϸ, "������ϸ�ϴ�")
    With rs��ϸ
        If .RecordCount = 0 Then
            ShowMsgbox "û����ص���ϸ��¼,������Щ��Ŀδ����ҽ������!"
            Exit Function
        End If
        Do While Not .EOF
            If Nvl(!��Ŀ����) = "" Then
                ShowMsgbox "����ϸ�д�����ص�ҽ����Ŀ"
                Exit Function
            End If
            .MoveNext
        Loop
        .MoveFirst
        lng����ID = 0
        str������¼�� = ""
        Dim strժҪ As String
        
        Do While Not .EOF
            If lng����ID <> Nvl(!����ID, 0) Then
                 '������һ�ŵ���
                 'AZYH    PCHAR   סԺ��
                 'ACFRQ   PCHAR   �������ڣ�YYYY-MM-DD��
                 'AYS PCHAR   ҽ��
                 'AKS PCHAR   ����
                 strInPut = Lpad(Nvl(!סԺ��, 0), 8)
                 strInPut = strInPut & "||" & Nvl(!�Ǽ�ʱ��)
                 strInPut = strInPut & "||" & Nvl(!ҽ��)
                 strInPut = strInPut & "||" & Nvl(!��������)
                 If ҵ������_�ɶ�����(���Ӵ�������, strInPut, strOutPut) = False Then Exit Function
                 str������¼�� = strOutPut
                 
                 '������������
                'AZYH    PCHAR   סԺ��
                'ACFID   PCHAR   ������¼�ţ�ͨ������ADDCFJL���ص�ֵ��
                 strInPut = Lpad(Nvl(!סԺ��, 0), 8)
                 strInPut = strInPut & "||" & str������¼��
                 If ҵ������_�ɶ�����(������������, strInPut, strOutPut) = False Then
                    '��ɾ�����ŵ���
                    Call ҵ������_�ɶ�����(ɾ���������ݼ�����ϸ, str������¼��, strOutPut)
                    Exit Function
                 End If
            End If
            '���Ӵ�����ϸ
            'ACFID   PCHAR   ������¼��
            'AYPBH   PCHAR   ҩƷ���(�籣ҩƷ���)
            'ASL PCHAR   ����(����Ϊ����)
            'ADJ PCHAR   ����
            strInPut = str������¼��
            strInPut = strInPut & "||" & Nvl(!��Ŀ����)
            strInPut = strInPut & "||" & Nvl(!����)
            strInPut = strInPut & "||" & Nvl(!�۸�)
            
            If ҵ������_�ɶ�����(���Ӵ�����ϸ, strInPut, strOutPut) = False Then
                '��ɾ�����ŵ���
                Call ҵ������_�ɶ�����(ɾ���������ݼ�����ϸ, str������¼��, strOutPut)
                Exit Function
            End If
           '������ϸ��¼��||�Էѱ���||�Էѽ��
           'ժҪ����ֵ:������¼��||��ϸ��¼��||�Էѱ���||�Էѽ��||סԺ��
            strժҪ = str������¼�� & "||" & strOutPut & "||" & Nvl(!סԺ��)
            '�����ϴ���־
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strժҪ & "')"
            zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
            .MoveNext
        Loop
    End With
    �����Ǽ�_�ɶ����� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function Readģ������(ByVal intҵ������ As ҵ������_�ɶ�����, ByVal strInputString As String, ByRef strOutPutstring As String)
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
                    strArr = Split(strText, vbTab)
                    If Val(strArr(0)) = 1 Then
                            str = strArr(1)
                            Exit Do
                    End If
                Else
                        If blnStart Then
                            If strText = "" Then
                                strText = "" & vbTab
                            End If
                            strArr = Split(strText, vbTab)
                            
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
    If InStr(1, strOutPutstring, "@$") <> 0 Then
        strOutPutstring = Split(strOutPutstring, "@$")(1)
    End If
    Exit Function
ErrHand:
    DebugTool Err.Description
    Exit Function
End Function
Private Sub OpenRecordset_�ɶ�����(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSql As String = "")
'���ܣ��򿪼�¼��
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSql = "", gstrSQL, strSql))
    rsTemp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle_�ɶ�����, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Public Function סԺ�������_�ɶ�����(rsExse As Recordset, ByVal lng����ID As Long, Optional bln���ʴ� As Boolean = True) As String
    
    'rsExse:�ַ���
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim rsTemp As New ADODB.Recordset
    
    Dim lng��ҳid As Long
    Dim strInPut As String, strOutPut   As String
    Dim strArr As Variant
    Dim strסԺ�� As String, str���㷽ʽ As String
    
    Dim intMouse As Integer
    
    Err = 0: On Error GoTo ErrHand:
    g�������_�ɶ�����.����ID = 0
    If rsExse.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "�������")
    If IsNull(rsTemp("��ҳID")) = True Then
        MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    lng��ҳid = rsTemp("��ҳID")
    
    gstrSQL = "Select ҽ��סԺ�� סԺ�� From �����ʻ� where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡסԺ��"
    If rsTemp.EOF Then
        ShowMsgbox "�ò��˲���ҽ������!"
        Exit Function
    End If
    strסԺ�� = Lpad(Nvl(rsTemp!סԺ��), 8)
    
    Screen.MousePointer = vbHourglass
    If ����סԺ��ϸ��¼(lng����ID, lng��ҳid) = False Then Exit Function
    'AZYH    PCHAR   סԺ��
    'ISPREV  PCHAR   Ԥ�����־��'0'����ʾԤ���㣩
    strInPut = strסԺ��
    strInPut = strInPut & "||0"
    If ҵ������_�ɶ�����(��Ժ����, strInPut, strOutPut) = False Then Exit Function
    strArr = Split(strOutPut, "||")
    
    '����ֵ
    'Ӧ֧��ͳ���||���ⶥ�Ը�||�����Ը�С��||����֧���ϼ�||ͳ��֧��ͳ���||����Ӧ��֧��||���β����㸶��||���β����������||ʵ�ʿۼ�����||ͳ��ⶥ���||
    'ͳ���𸶽��||סԺ��¼��||���˼�¼��||����||סԺ��||���ֱ��||��������||����||ҽ�ƻ�����||��Ժ����||��Ժ����||�ѽ���ͳ���||����ҽ�Ʒ�С��||����ҩƷ��||
    '��������||�������Ʒ�||����������||�Ը�С��||�Ը�ҩƷ��||�Ը�����||�Ը����Ʒ�||�Ը�������||����ҩƷ��||��������||�������Ʒ�||����������||ͳ��֧������||
    'ͳ�����Ʒ�||ͳ��������||��Ժ��־||�����־||�����־||����ҽ��״̬||���㷽ʽ||��˷�ʽ||��������||��λ���||��λ����||��ᱣ�Ϻ�||����||�Ա�||��������||Ԥ�ɽ��||
    '����Ӧ�����||����ʵ�����||�˿���||����ʵ��֧�����||�籣������||�����������||��������־||���ջ�����||����Ա���||������ȡʱ��||ҽ�Ʊ��ձ��||�籣��������||
    '������������||�������ܱ�־||�����𸶿ۼ���־||����������||����������||�����㸶����||�������㸶��||�����㸶�����||����������ܿ�ʼ����||�㸶�ܶ�||
    With g�������
        .ҽ������ = Val(strArr(4))
        .�����㸶�� = Val(strArr(6))
    End With
    
    str���㷽ʽ = "ҽ������;" & g�������.ҽ������ & ";0"
    str���㷽ʽ = str���㷽ʽ & "|�����㸶��;" & g�������.�����㸶�� & ";0"
    סԺ�������_�ɶ����� = str���㷽ʽ
    g�������_�ɶ�����.����ID = lng����ID   '��ʾ�ò����Ѿ��������������
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function ����סԺ��ϸ��¼(ByVal lng����ID As Long, ByVal lng��ҳid As Long) As Boolean
    '���������ϸ��¼
    Dim cnTemp As New ADODB.Connection
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim strInPut  As String, strOutPut As String
    Dim strArr
    Dim strסԺ�� As String, str������¼�� As String
    
    Err = 0
    On Error GoTo ErrHand:
      
    
    Call DebugTool("��������")
    cnTemp.ConnectionString = gcnOracle.ConnectionString
    cnTemp.Open
    Call DebugTool("�����ӳɹ�����ʼ�����ϸ���ݵĺϷ��ԡ�")
    
      
      
      
    ����סԺ��ϸ��¼ = False
    
    '����δ�ϴ���ϸ�������Ա����ϴ�����ϸ�����ϴ�����ϸ��
    gstrSQL = "Select A.ID,A.NO,A.��¼����,A.��¼״̬,A.���,A.����ID,A.��ҳID,to_char(A.����ʱ��,'yyyy-mm-dd')  as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս��" & _
              "         ,M.���� as ��������,A.�շ�ϸĿID,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
              "         ,C.��Ŀ����,C.��ע,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,A.����Ա����,B.���㵥λ" & _
              "  From ���˷��ü�¼ A,�շ�ϸĿ B,(Select * From ����֧����Ŀ where ����=" & gintInsure & ") C,������ҳ D,���ű� M" & _
              "  where A.����ID=" & lng����ID & " and A.��ҳID=" & lng��ҳid & " and A.���ʷ���=1 and A.ʵ�ս��<>0 and nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 " & _
              "        and A.��������id =M.id(+) and A.����ID=D.����ID and A.��ҳID=D.��ҳID And D.����=" & gintInsure & _
              "        and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID(+) " & _
              "  Order by A.����ID,A.��¼����,A.No,A.��¼״̬,A.���"
    Call OpenRecordset(rs��ϸ, "�������")
    
    gstrSQL = "Select ҽ��סԺ�� סԺ�� From �����ʻ� where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡסԺ��"
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "�ڱ����ʻ��в����ڸò���"
        Exit Function
    End If
    strסԺ�� = Nvl(rsTemp!סԺ��, 0)
    
   With rs��ϸ
'        If .RecordCount = 0 Then
'            ShowMsgbox "û����ص���ϸ��¼,������Щ��Ŀδ����ҽ������!"
'            Exit Function
'        End If
        Do While Not .EOF
            If Nvl(!��Ŀ����) = "" Then
                ShowMsgbox "����ϸ�д�����ص�ҽ����Ŀ"
                Exit Function
            End If
            .MoveNext
        Loop
        If Not .EOF Then .MoveFirst
        Dim strNO As String
        
        str������¼�� = ""
        strNO = ""
        Dim strժҪ As String
        
        Do While Not .EOF
            If strNO <> Nvl(!��¼����, 0) & "_" & Nvl(!NO) & "_" & Nvl(!��¼״̬, 0) Then
                strNO = Nvl(!��¼����, 0) & "_" & Nvl(!NO) & "_" & Nvl(!��¼״̬, 0)
                 '������һ�ŵ���
                 'AZYH    PCHAR   סԺ��
                 'ACFRQ   PCHAR   �������ڣ�YYYY-MM-DD��
                 'AYS PCHAR   ҽ��
                 'AKS PCHAR   ����
                 
                 strInPut = Lpad(strסԺ��, 8)
                 strInPut = strInPut & "||" & Nvl(!�Ǽ�ʱ��)
                 strInPut = strInPut & "||" & Nvl(!ҽ��)
                 strInPut = strInPut & "||" & Nvl(!��������)
                 If ҵ������_�ɶ�����(���Ӵ�������, strInPut, strOutPut) = False Then Exit Function
                 str������¼�� = strOutPut
                 
                 '������������
                'AZYH    PCHAR   סԺ��
                'ACFID   PCHAR   ������¼�ţ�ͨ������ADDCFJL���ص�ֵ��
                 strInPut = Lpad(strסԺ��, 8)
                 strInPut = strInPut & "||" & str������¼��
                 If ҵ������_�ɶ�����(������������, strInPut, strOutPut) = False Then
                    '��ɾ�����ŵ���
                    Call ҵ������_�ɶ�����(ɾ���������ݼ�����ϸ, str������¼��, strOutPut)
                    Exit Function
                 End If
            End If
            '���Ӵ�����ϸ
            'ACFID   PCHAR   ������¼��
            'AYPBH   PCHAR   ҩƷ���(�籣ҩƷ���)
            'ASL PCHAR   ����(����Ϊ����)
            'ADJ PCHAR   ����
            strInPut = str������¼��
            strInPut = strInPut & "||" & Nvl(!��Ŀ����)
            strInPut = strInPut & "||" & Nvl(!����)
            strInPut = strInPut & "||" & Nvl(!�۸�)
            
            If ҵ������_�ɶ�����(���Ӵ�����ϸ, strInPut, strOutPut) = False Then
                '��ɾ�����ŵ���
                Call ҵ������_�ɶ�����(ɾ���������ݼ�����ϸ, str������¼��, strOutPut)
                Exit Function
            End If
           '������ϸ��¼��||�Էѱ���||�Էѽ��
           'ժҪ����ֵ:������¼��||��ϸ��¼��||�Էѱ���||�Էѽ��||סԺ��
            strժҪ = str������¼�� & "||" & strOutPut & "||" & strסԺ��
            '�����ϴ���־
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strժҪ & "')"
             cnTemp.Execute gstrSQL, , adCmdStoredProc
            .MoveNext
        Loop
    End With
    ����סԺ��ϸ��¼ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function ҽ������_�ɶ�����() As Boolean
    ҽ������_�ɶ����� = frmSet�ɶ�����.��������
End Function
