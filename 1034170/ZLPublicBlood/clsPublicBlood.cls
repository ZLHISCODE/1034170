VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicBlood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mfrmBloodPeoPle As frmBloodPeoPle
Attribute mfrmBloodPeoPle.VB_VarHelpID = -1

Public Event PeopleChange(ByVal strRetrun As String)  'Ҫ�����������״̬������
Public Event AfterPatiFind(ByVal strIDKindstr As String, ByVal strValue As String)  '���ҵ�IDKindStr���濨Ƭ�ϣ��򷵻��¼��е���������

'ҽ������ö��
Public Enum Enum_Advice
    Advice_�¿� = 0
    Advice_�޸�
    Advice_ɾ��
    Advice_У��
    Advice_����
    Advice_����
    Advice_����
    Advice_�������� 'סԺר��
End Enum

Public Enum Enum_CheckType
    ��Ѫ�˶� = 0
    ���պ˶�
    ִ�к˶�
End Enum

'���������������Ƿ����Ե��û�����Ѫ��ϵͳ���������HIS����������ٴ���������sp�汾�����������Դ���
'Ŀǰ������֧�ֵ��Ե�汾��34.110,35.20   ˵������ʽ�汾���ø�����
Public Property Get IsPilotBlood() As Boolean
    IsPilotBlood = True
End Property

Public Function zlInitCommon(ByVal cnOracle As ADODB.Connection, Optional ByVal strDBUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż����ݿ�����
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-��ǰ���ݿ��¼�û���
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set gcnOracle = cnOracle: gstrDBUser = strDBUser
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrUnitName = gobjComlib.GetUnitName
    Call InitLocPar
    
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub ZlSetFormSize(ByVal frmParent As Object, ByVal lngMinWidth As Long, ByVal lngMinHeight As Long, ByVal lngMaxWidth As Long, ByVal lngMaxHeight As Long, ByRef lngWindowLong As Long, Optional ByVal blnUnLoad As Boolean = False)
'���ܣ����ô����С��Χ
'˵����������ػ�ȡlngWindowLong������ж�ش���lngWindowLong
    Call SetFormSize(frmParent, lngMinWidth, lngMinHeight, lngMaxWidth, lngMaxHeight, lngWindowLong, blnUnLoad)
End Sub

Public Function zlBloodUserIdentity(ByVal frmParent As Object, ByVal lngModul As Long, ByVal lngReviewDeptID As Long, ByVal lngCheckDeptId As Long, ByVal str��Сʱ�� As String, ByVal str���ʱ�� As String, _
    Optional ByVal blnSelectCheckUser As Boolean = True, Optional ByVal intMode As Enum_CheckType = ��Ѫ�˶�, Optional ByVal strIDs As String = "") As Boolean
'����:ѪҺ��Ѫ/����/ִ��������֤
'��Σ�
'       lngModul=����ģ��
'       lngReviewDeptID :�������ID(ȷ��������Ա) User2
'       lngCheckDeptId���˶Կ���ID(ȷ���˶���Ա) User1
'       str��Сʱ��-str���ʱ�䣺ȷ���˶�ʱ��ķ�Χ
'       blnSelectCheckUser���˶����Ƿ���Ҫ������֤(������ǿ����Ҫ����������֤)
'       strIDs�����ڶ��ѪҺ���з�Ѫ�����ա�ִ�б��봫���շ�ID��(���ŷָ�)����������ÿһ��ѪҺ�ĺ˶Խ��
    zlBloodUserIdentity = frmUserCheck.ShowMe(frmParent, lngModul, lngReviewDeptID, lngCheckDeptId, str��Сʱ��, str���ʱ��, blnSelectCheckUser, intMode, strIDs)
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'ҽ���������
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'��ʾ�䷢��Ϣ
Public Function zlGetBloodListInfo() As Object
    Set zlGetBloodListInfo = New frmBloodListInfo
End Function

'��Ѫִ��
Public Function zlGetBloodExec() As Object
    Set zlGetBloodExec = New frmBloodExec
End Function


Public Function ShowBloodInstantRptPrint(ByVal objfrm As Object, ByVal lngActiveID As Long) As Boolean
'���ܣ�����ʿվ��ҽ������վ����(��Ѫִ�е���ӡ)
'������ objFrm--����������
'           lngActiveID--ҽ��ID
    ShowBloodInstantRptPrint = zlBloodInstantRptPrint(objfrm, lngActiveID)
End Function

Public Function AdviceOperation(ByVal lngMoudle As Long, ByVal lngҽ��ID As Long, ByVal intOperation As Enum_Advice, Optional ByVal blnMoved As Boolean = False, _
        Optional ByRef strErrInfo As String = "") As Boolean
'���ܣ�ҽ���������ýӿڣ��¿���ɾ�������͡�����ʱ�˷����ĵ��������ҽ�����������е��ã��޸ġ�У�ԡ�����Ϊ����У���飬��������֮ǰ��
'���:
'       lngMoudle:����ģ���
'       lngҽ��ID:ѪҺҽ����ҽ��ID
'       intOperation:ҽ����������(ö��),�����¿����޸ġ�ɾ����У�ԡ����ϡ����͡�����
'       blnMoved:������ʷ�����Ƿ�ת��
'���Σ�
'       strErrInfo���ӿڷ���FALSEʱ����Ϣ
'���أ��ɹ�=TRUE��ʧ��=False
    AdviceOperation = zlAdviceOperation(lngMoudle, lngҽ��ID, intOperation, blnMoved, strErrInfo)
End Function

Public Function GetPrepareBloodRs(ByVal lngҽ��ID As Long, rsBlood As ADODB.Recordset) As Boolean
'����:��ȡ��Ѫ��¼�����Ϣ���Լ�¼������
'��Σ�lngҽ��ID����Ѫҽ��ID
'���Σ�rsTemp��ѪҺ��Ѫ��¼��¼��
'���أ�TRUE=�ɹ�,False-ʧ��
    Dim strSQL As String
    On Error GoTo ErrHand
    strSQL = "Select Id, ������Դ, ����id, ��ҳid, ����id, ִ�в���id, ��¼״̬, ������, ����ʱ��, �����, ���ʱ��, ��¼����" & vbNewLine & _
        "From ѪҺ��Ѫ��¼" & vbNewLine & _
        "Where ����id = [1]"
    Set rsBlood = gobjDatabase.OpenSQLRecord(strSQL, "GetPrepareBloodRs", lngҽ��ID)
    GetPrepareBloodRs = Not rsBlood.EOF
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function AdviceTermination(ByVal lngMoudle As Long, ByVal lngҽ��ID As Long, Optional ByVal blnCance As Boolean = False, Optional ByVal blnMoved As Boolean = False, _
    Optional ByRef strErrInfo As String = "", Optional ByRef blnFallback As Boolean = False) As Boolean
'���ܣ�ҽ�����δ�õ���(��Ѫ��ֹ���ܴ���)
'���:
'       lngMoudle:����ģ���
'       lngҽ��ID:ѪҺҽ����ҽ��ID
'       blnCance:TRUE:ȡ����Ѫ��ֹ,False:��Ѫ��ֹ
'       blnMoved:������ʷ�����Ƿ�ת��
'���Σ�
'       strErrInfo���ӿڷ���FALSEʱ����Ϣ
'       blnFallback:���δ�õ�ͬʱ�Ƿ����ҽ��
'���أ��ɹ�=TRUE��ʧ��=False
    AdviceTermination = True
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'���˿�Ƭ�б����
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function zlGetBloodPeople() As Object
    Set zlGetBloodPeople = mfrmBloodPeoPle
End Function

Public Function zlRefreshBloodPeople(ByVal frmParent As Object, ByVal strRule As String, ByVal rs As ADODB.Recordset, ByVal lngMode As Long, Optional ByVal imgList As Object, Optional ByVal strIDKindstr As String = "") As Boolean
    Call mfrmBloodPeoPle.UserInit(frmParent, strRule, imgList, lngMode, strIDKindstr)
    Call mfrmBloodPeoPle.ShowPeople(rs)
    zlRefreshBloodPeople = True
End Function

Public Sub BloodPeopleLocation(strTitle As String, strfind As String)
    '��λ��ָ������Ա����
    Call mfrmBloodPeoPle.SetCardFocus(strTitle, strfind)
End Sub

Private Sub mfrmBloodPeoPle_AfterPatiFind(ByVal strIDKindstr As String, ByVal strValue As String)
    RaiseEvent AfterPatiFind(strIDKindstr, strValue)
End Sub

Private Sub mfrmBloodPeoPle_CardChanged()
    RaiseEvent PeopleChange(mfrmBloodPeoPle.strReturn)
End Sub

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'��Ѫ��Ӧ�Ǽ�
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function zlShowBloodReaction(ByVal frmParent As Object, ByVal lngSys As Long, ByVal lngMoudle As Long, lng���� As Long, lng����ID As Long, lng��ҳid As Long, _
                    lng������Դ As Long, Optional ByVal lngisMoudal As Long = 1, Optional ByVal lng�շ�id As Long = 0)
'-------------------------------------------------------------------------------------------------------------------------
'���ܣ���ָ�����˵���Ѫ��Ӧ���в鿴��༭
'��Σ�frmParent-���ô������
'          lngMoudle-����ģ���,lngSys--����ϵͳ��
'         lng����-1:ҽ��վ����;2-��Ѫ�Ƶ���
'         lng����id������ID
'         lng��ҳid����ҳID�������(סԺ����Ϊ��ҳID�����ﲡ��Ϊ�����)
'         lng������Դ��1-����;2-סԺ
'         lngisMoudal:1-ģ̬��0-��ģ̬,Ĭ��Ϊģ̬����
'-------------------------------------------------------------------------------------------------------------------------
    Dim strPrivs As String
    strPrivs = GetInsidePrivs(2200, p��Ѫ��Ӧ����)
    frmBloodReaction.BloodReaction frmParent, lng����, lng����ID, lng��ҳid, lng������Դ, lngSys, lngMoudle, strPrivs, lngisMoudal, lng�շ�id
End Function

Public Function zlShowBloodReactionRecord(ByVal frmParent As Variant, ByVal lngSys As Long, ByVal lngMoudle As Long, lng���� As Long, Optional ByVal lngisMoudal As Long = 1)
'-------------------------------------------------------------------------------------------------------------------------
'���ܣ��Բ���Ա���������²��˵���Ѫ��Ӧ���в鿴��༭
'��Σ�frmParent-���ô������, BH���ô��븸����HWND
'          lngMoudle-����ģ��ţ�lngsys-����ϵͳ��
'          lng����-  0:����ҽ��վ����1:סԺҽ��վ����;2-��Ѫ�Ƶ���
'          lngisMoudal:1-ģ̬��0-��ģ̬
'-------------------------------------------------------------------------------------------------------------------------
    Dim strPrivs As String
    strPrivs = GetInsidePrivs(2200, p��Ѫ��Ӧ����)
    frmBloodReactionRecord.BloodReactionRecord frmParent, lng����, lngSys, lngMoudle, strPrivs, lngisMoudal
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'ѪҺ���յǼ�
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function zlShowBloodReceives(ByVal frmParent As Variant, ByVal lngSys As Long, ByVal lngMoudle As Long, Optional ByVal lngisMoudal As Long = 1, _
            Optional int���� As Integer = 0)
'���ܣ��Բ���Ա���������²��˵�ѪҺ���н��ղ���(CodeMan��ڵ���)
'��Σ�frmParent-���ô������, BH���ô��븸����HWND,lngisMoudal:1-ģ̬��0-��ģ̬
'          lngMoudle-����ģ���  lngSys:����ϵͳ��
'-------------------------------------------------------------------------------------------------------------------------
    Dim strPrivs As String
    strPrivs = GetInsidePrivs(100, pѪҺ���յǼ�)
    frmBloodReceivesRecord.BloodReceives frmParent, lngSys, lngMoudle, strPrivs, lngisMoudal, int����
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'��Ѫ���
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function zlGetBloodAuditState(ByVal intMode As Integer) As String
'���ܣ���ȡѪҺƷ���Ƿ�����Լ��Ƿ������Ѫ��
'��Σ�intMode-ģʽ��1-����������˵�ѪҺ��Ʒ��id,2-�������������ѪҺ��Ʒ��id
'
'-------------------------------------------------------------------------------------------------------------------------
    zlGetBloodAuditState = GetBloodAuditState(intMode)
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'ҽ��վ��Ϣ״̬
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function zlIsBloodMessageDone(ByVal intMode As Integer, ByVal lng����ID As Long, ByVal lng����id As Long, ByVal int�Ķ����� As Integer, ByVal lng�Ķ�����id As Long) As Boolean
'���ܣ�����ҽ��վ�ò��˵���Ϣ״̬
'��Σ�intMode-ģʽ��1-Ѫ�������Ϣ,2-��Ѫ��Ӧ��Ϣ
'
'-------------------------------------------------------------------------------------------------------------------------
    zlIsBloodMessageDone = IsBloodMessageDone(intMode, lng����ID, lng����id, int�Ķ�����, lng�Ķ�����id)
End Function
Private Sub Class_initialize()
    Set mfrmBloodPeoPle = New frmBloodPeoPle
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not mfrmBloodPeoPle Is Nothing Then
        Unload mfrmBloodPeoPle
        Set mfrmBloodPeoPle = Nothing
    End If
    
    If Not gobjPublicExpense Is Nothing Then
        Set gobjPublicExpense = Nothing
    End If
    Call CloseSquareCardObject
    If Err <> 0 Then Err.Clear
End Sub