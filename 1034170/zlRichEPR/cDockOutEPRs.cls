VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockOutEPRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�̳д����¼�
Public Event RequestRefresh() 'Ҫ��������ˢ��
Public Event Activate()
Public Event ClickDiagRef(DiagnosisID As Long, Modal As Byte)

Private WithEvents frmOutEPRs As frmDockOutEPRs
Attribute frmOutEPRs.VB_VarHelpID = -1

Private Sub Class_initialize()
    Set frmOutEPRs = New frmDockOutEPRs
End Sub

Private Sub Class_Terminate()
    Unload frmOutEPRs
    Set frmOutEPRs = Nothing
End Sub


Private Sub frmOutEPRs_RequestRefresh()
    RaiseEvent RequestRefresh
End Sub

Private Sub frmOutEPRs_Activate()
    RaiseEvent Activate
End Sub

Private Sub frmOutEPRs_ClickDiagRef(DiagnosisID As Long, Modal As Byte)
    RaiseEvent ClickDiagRef(DiagnosisID, Modal)
End Sub
Public Sub SetFontSize(ByVal bytSize As Byte)
    Call frmOutEPRs.SetFontSize(bytSize)
End Sub
Public Function zlGetForm() As Object
    Set zlGetForm = frmOutEPRs
End Function

Public Sub zlDefCommandBars(ByVal CommandBars As Object)
    Call frmOutEPRs.zlDefCommandBars(CommandBars)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmOutEPRs.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmOutEPRs.zlExecuteCommandBars(Control)
End Sub

Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngBillId As Long, ByVal lngDeptId As Long, Optional ByVal bnEdit As Boolean, _
                            Optional ByVal blnMoved As Boolean, Optional ByVal blnForce As Boolean, Optional ByVal lngAdviceID As Long) As Long
    '����:����ˢ��ָ�����˵Ĳ������ݣ�����������ṩ�༭����
    '����:  lngPatiId-����id;
    '       lngBillId-�Һ�id;
    '       lngDeptId-��ǰ�������ţ�ע�ⲻ�ǲ��˱��ξ�����ң�
    '       blnEdit-�Ƿ������༭��ͨ����ǰ�������Ų��ǲ��˱��ξ�����ң���Ӧ�ò������༭��
    '       blnMoved-�����Ƿ�ת��
    '       lngAdviceID ҽ��ID��Ŀǰֻ������ģ����ô���
    zlRefresh = frmOutEPRs.zlRefresh(lngPatiID, lngBillId, lngDeptId, bnEdit, blnForce, blnMoved, lngAdviceID)
End Function

Public Function zlOpenDefaultEPR(ByVal str�Һŵ��� As String) As Boolean

    '����:�Զ�����������
    '����:  str�Һŵ���-�Һŵ���;
    
    Dim strSQL As String
    Dim bytKind As Byte
    Dim rs As New ADODB.Recordset
        
    strSQL = "Select Nvl(����,0) As ����,Nvl(����,0) As ���� From ���˹Һż�¼ Where No=[1] and ��¼����=1  and ��¼״̬=1"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "���ﲡ��", str�Һŵ���)
    If rs.BOF = False Then
        If rs("����").Value = 1 Then
            bytKind = 3
        ElseIf rs("����").Value = 1 Then
            bytKind = 2
        Else
            bytKind = 1
        End If
        
        zlOpenDefaultEPR = frmOutEPRs.zlOpenDefaultEPR(bytKind)
        
    End If
    
End Function