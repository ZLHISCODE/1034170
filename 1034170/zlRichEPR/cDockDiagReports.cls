VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockDiagReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�̳д����¼�
Public Event Activate()

Private WithEvents frmInDiagReport As frmDockDiagReports
Attribute frmInDiagReport.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set frmInDiagReport = New frmDockDiagReports
End Sub

Private Sub Class_Terminate()
    Set frmInDiagReport = Nothing
End Sub

Private Sub frmInTend_Activate()
    RaiseEvent Activate
End Sub

Public Function zlGetForm() As Object
    Set zlGetForm = frmInDiagReport
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, _
                            ByVal cbsMain As Object, _
                            ByVal int������Դ As Integer, _
                            Optional ByVal bln��ʿվ As Boolean)
    Call frmInDiagReport.zlDefCommandBars(frmParent, cbsMain, int������Դ, bln��ʿվ)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInDiagReport.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInDiagReport.zlExecuteCommandBars(Control)
End Sub

Public Function zlRefresh(ByVal lngPatiID As Long, _
                            ByVal strPageID As String, _
                            ByVal lngDeptId As Long, _
                            ByVal int���� As Integer, _
                            Optional ByVal bnEdit As Boolean, _
                            Optional ByVal bln����ת�� As Boolean) As Long
    '����:����ˢ��ָ�����˵Ĳ������ݣ�����������ṩ�༭����
    '����:  lngPatiId-����id;
    '       lngPageId-��ҳid;
    '       lngDeptId-��ǰ�������ţ�ע�ⲻ�ǲ��˵�ǰ���ڲ������Ժ������
    '       blnEdit-�Ƿ������༭��ͨ����ǰ�������Ų��ǲ������ڲ������Ժ��������Ӧ�ò������༭��ҽ��վֱ������Ϊ�������༭
    zlRefresh = frmInDiagReport.zlRefresh(lngPatiID, strPageID, lngDeptId, bnEdit, int����, bln����ת��)
    
End Function

