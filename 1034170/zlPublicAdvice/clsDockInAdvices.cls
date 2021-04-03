VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockInAdvices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event Activate() '�Ӵ��弤��
Public Event RequestRefresh(ByVal RefreshNotify As Boolean) 'Ҫ��������ˢ��
Public Event StatusTextUpdate(ByVal Text As String) 'Ҫ�����������״̬������
Public Event ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean) 'Ҫ��鿴����
Public Event PrintEPRReport(ByVal ����ID As Long, ByVal Preview As Boolean) 'Ҫ���ӡ����
Public Event ViewPACSImage(ByVal ҽ��ID As Long) 'Ҫ����й�Ƭ
Public Event ExecLogNew(ByVal ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal ����ID As Long, ��� As Boolean) 'ִ������Ǽ�
Public Event ExecLogModi(ByVal ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal ����ID As Long, ByVal ִ��ʱ�� As String, ��� As Boolean) 'ִ������޸�
Public Event EditDiagnose(ParentForm As Object, ByVal ����ID As Long, ByVal ��ҳID As Long, ByVal ����ID As Long, ByVal str���� As String, Succeed As Boolean) '�༭סԺ���
Public Event SetEditState(ByVal blnEditState As Boolean)    '�༭״̬ʱ���ò˵��Ϳ�ת�ƽ���Ĺ���
Public Event DoByAdvice(ByVal lngҽ��ID As Long, ByVal lng���ID As Long, ByVal lngWayID As Long, ByVal strTag As String)

Private WithEvents mclsInAdvice As zlCISKernel.clsDockInAdvices
Attribute mclsInAdvice.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mclsInAdvice = New zlCISKernel.clsDockInAdvices
End Sub

Private Sub Class_Terminate()
    Set mclsInAdvice = Nothing
End Sub

Private Sub mclsInAdvice_Activate()
    RaiseEvent Activate
End Sub

Private Sub mclsInAdvice_SetEditState(ByVal blnEditState As Boolean)
    RaiseEvent SetEditState(blnEditState)
End Sub

Private Sub mclsInAdvice_RequestRefresh(ByVal RefreshNotify As Boolean)
    RaiseEvent RequestRefresh(RefreshNotify)
End Sub

Private Sub mclsInAdvice_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Private Sub mclsInAdvice_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
    RaiseEvent ViewEPRReport(����ID, CanPrint)
End Sub

Private Sub mclsInAdvice_PrintEPRReport(ByVal ����ID As Long, ByVal Preview As Boolean)
    RaiseEvent PrintEPRReport(����ID, Preview)
End Sub

Private Sub mclsInAdvice_ViewPACSImage(ByVal ҽ��ID As Long)
    RaiseEvent ViewPACSImage(ҽ��ID)
End Sub

Private Sub mclsInAdvice_ExecLogNew(ByVal ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal ����ID As Long, ��� As Boolean)
    RaiseEvent ExecLogNew(ҽ��ID, ���ͺ�, ����ID, ���)
End Sub

Private Sub mclsInAdvice_ExecLogModi(ByVal ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal ����ID As Long, ByVal ִ��ʱ�� As String, ��� As Boolean)
    RaiseEvent ExecLogModi(ҽ��ID, ���ͺ�, ����ID, ִ��ʱ��, ���)
End Sub

Private Sub mclsInAdvice_EditDiagnose(ParentForm As Object, ByVal ����ID As Long, ByVal ��ҳID As Long, ByVal ����ID As Long, ByVal str���� As String, Succeed As Boolean)
    RaiseEvent EditDiagnose(ParentForm, ����ID, ��ҳID, ����ID, str����, Succeed)
End Sub

Private Sub mclsInAdvice_DoByAdvice(ByVal lngҽ��ID As Long, ByVal lng���ID As Long, ByVal lngWayID As Long, ByVal strTag As String)
    RaiseEvent DoByAdvice(lngҽ��ID, lng���ID, lngWayID, strTag)
End Sub

Public Function zlGetForm() As Object
'���ܣ���ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ
    '��ʼ��ҽ������:�ɷ�ɢ�������GetForm��ȥ��ʼ��
    Set zlGetForm = mclsInAdvice.zlGetForm
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, ByVal int���� As Integer, _
    Optional ByVal blnInsideTools As Boolean = False, Optional ByRef objSquareCard As Object, Optional ByVal blnModalNew As Boolean)
'���ܣ������Ӵ���Ĳ˵��͹�����(����������Ҫʹ�õĲ˵��͹�����)
'������int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      CommandBars=�����ڲ鿴ʱ���Բ���(����Nothing)
'      blnInsideTools=�Ƿ񽫰�ť��ʾ�������ڲ��������ϣ�����ʿվ����ʱ��Ч��
'      objSquareCard=ͨ��ҽ������վ�����һ��ͨ���㲿��(סԺҽ��վ�ڷ���ʱ�ٴ�������ʼ��)

    Call mclsInAdvice.zlDefCommandBars(frmParent, CommandBars, int����, blnInsideTools, objSquareCard, blnModalNew)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mclsInAdvice.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mclsInAdvice.zlExecuteCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBarsDirect(ByVal Control As CommandBarControl, ByRef frmParent As Object, ByRef strPrivs As String, _
    ByVal bln���� As Boolean, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal bytӤ�� As Byte, _
    ByVal lng����ID As Long, ByVal lng����id As Long, ByVal lngǰ��ID As Long, ByVal lng�������ID As Long, ByVal int���� As Integer, _
    ParamArray arrPar() As Variant)
'���ܣ��ṩ��������ҽ�������Ľӿ�
    Call mclsInAdvice.zlExecuteCommandBarsDirect(Control, frmParent, strPrivs, bln����, lng����ID, lng��ҳID, bytӤ��, _
        lng����ID, lng����id, lngǰ��ID, lng�������ID, int����, arrPar)
End Sub

Public Sub zlCheckPrivs(ByVal Control As CommandBarControl, ByVal int���� As Integer)
'���ܣ����˵���ť��Ȩ�ޣ���������ɼ���
    Call mclsInAdvice.zlCheckPrivs(Control, int����)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mclsInAdvice.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal lng����id As Long, _
    ByVal int״̬ As Integer, Optional ByVal blnMoved As Boolean, Optional ByVal lngǰ��ID As Long, _
    Optional ByVal intִ��״̬ As Integer, Optional ByVal lng�������ID As Long, Optional ByVal lng·��״̬ As Long = -1, _
    Optional ByVal lngҽ������ID As Long, Optional ByRef objMip As Object, Optional ByVal intӤ�� As Integer = -1, Optional ByVal lng����ҽ��ID As Long)
'���ܣ�ˢ��סԺҽ������
'������int״̬=0-��Ժ����,1-Ԥ��Ժ����,2-��Ժ����,3-�����ﲡ��(��Ժ),4-�ѻ��ﲡ��,5-���ת�Ʋ���(��Ժ),6-ת�ƻ�ת��������ס����(������Ժ����ס)
'      lng����ID��lng����ID=����5-���ת�Ʋ��ˡ�ʱΪ����ԭ������ԭ����
'      lngǰ��ID=����ҽ��վ����ʱ����
'      intִ��״̬=����ҽ��վ����ʱ����,ҽ��ִ����Ŀ��ִ��״̬
'      lng�������ID=�����ǰҽ��վ�ǻ��ﲡ�ˣ���Ϊ�������ID�������ҽ��վ���ã���Ϊҽ������ID
'      lng·��״̬=-1:δ����,0-�����ϵ���������1-ִ���У�2-����������3-�������
'      blnMoved=�ò��˵������Ƿ���ת��
'      lngҽ������ID=ҽ��վ����ID
'      strǰ��IDsҽ��վ���ڵ�ǰ����ִ�е�����ҽ��
'      objMip ��Ϣ���� zl9ComLib.clsMipModule
    Call mclsInAdvice.zlRefresh(lng����ID, lng��ҳID, lng����ID, lng����id, int״̬, blnMoved, lngǰ��ID, intִ��״̬, lng�������ID, lng·��״̬, lngҽ������ID, objMip, intӤ��, lng����ҽ��ID)
End Sub

Public Sub LocatedAdviceRow(ByVal lngҽ��ID As Long)
'���ܣ����ݹ���վ�����ҽ��id��ѡ�����Ҷ�λҽ��
    Call mclsInAdvice.LocatedAdviceRow(lngҽ��ID)
End Sub

Public Sub zlSeekAndViewEPRReport(ByVal lng����ID As Long)
'���ܣ���λ�������Ӧ��ҽ�������򿪱���鿴
    Call mclsInAdvice.zlSeekAndViewEPRReport(lng����ID)
End Sub

Public Sub zlAdvicePrint(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long)
'���ܣ���ӡҽ����
    Call mclsInAdvice.zlAdvicePrint(frmParent, lng����ID, lng��ҳID)
End Sub

Public Sub zlInitPath(ByRef clsDockPath As Object)
'���ܣ���ʼ���ٴ�·�����󣬴���ҽ��վ��ʿվ���õĶ���
    Call mclsInAdvice.zlInitPath(clsDockPath)
End Sub

Public Sub zlInitPlugIn(ByVal objPlugIn As Object)
'����:��������Ҳ������
'53740:������,2012-09-19
    Call mclsInAdvice.zlInitPlugIn(objPlugIn)
End Sub

Public Sub zlInitMip(ByRef objMip As Object)
'����:������Ϣ����
'������objMip zl9ComLib.clsMipModule
    Call mclsInAdvice.zlInitMip(objMip)
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'����:����ҽ���嵥�������С
'���:bytSize��0-С(ȱʡ)��1-��
    Call mclsInAdvice.SetFontSize(bytSize)
End Sub

Public Sub zlCompoundUnpack(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, Optional ByVal lngҽ������ID As Long, _
    Optional ByVal lngӤ������ID As Long, Optional ByVal lngӤ������ID As Long)
'���ܣ��������
    Call mclsInAdvice.zlCompoundUnpack(frmParent, lng����ID, lng����ID, lngҽ������ID, lngӤ������ID, lngӤ������ID)
End Sub

Public Function AdviceRisReport(frmParent As Object, ByVal lng����ID As Long) As Boolean
'���ܣ�������ӡRISԤԼ��
    AdviceRisReport = gobjKernel.AdviceRisReport(frmParent, lng����ID)
End Function