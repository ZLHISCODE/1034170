VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockInTends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�̳д����¼�
Public Event Activate()

Private WithEvents frmInTend As frmDockInTends
Attribute frmInTend.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set frmInTend = New frmDockInTends
End Sub

Private Sub Class_Terminate()
    Unload frmInTend
    Set frmInTend = Nothing
End Sub

Private Sub frmInTend_Activate()
    RaiseEvent Activate
End Sub

Public Function zlGetForm() As Object
    Set zlGetForm = frmInTend
End Function

Public Sub zlDefCommandBars(ByVal CommandBars As Object, Optional ByVal blnChildToolBar As Boolean = False)
    Call frmInTend.zlDefCommandBars(CommandBars, blnChildToolBar)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInTend.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInTend.zlExecuteCommandBars(Control)
End Sub

Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, ByVal bnEdit As Boolean, _
    Optional ByVal blnDoctorStation As Boolean, Optional ByVal lngUnitID As Long, Optional ByVal intState As Integer) As Long
    '����:����ˢ��ָ�����˵Ĳ������ݣ�����������ṩ�༭����
    '����:  lngPatiId-����id;
    '       lngPageId-��ҳid;
    '       lngDeptId-��ǰ�������ţ�ע�ⲻ�ǲ��˵�ǰ���ڲ������Ժ������
    '       lngUnitID��lngDeptId-��intState=��5-���ת�Ʋ��ˡ�ʱΪ����ԭ������ԭ����
    '       blnEdit-�Ƿ������༭��ͨ����ǰ�������Ų��ǲ������ڲ������Ժ��������Ӧ�ò������༭��ҽ��վֱ������Ϊ�������༭
    '       intState-��ʿվ��ҽ��վ����Ĳ���״̬TYPE_PATI_State
    zlRefresh = frmInTend.zlRefresh(lngPatiID, lngPageId, lngDeptId, bnEdit, , blnDoctorStation)
End Function

Public Sub zlLocateData(ByVal intType As Integer)
'����:��λ��������¼��Ż����������ݽ���
'����:intType=0-������¼����,1-������������
    Call frmInTend.zlLocateData(intType)
End Sub


Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '���������Ƿ��ṩ�ļ���Ϣ�б�
    '����:������
    '����:2012-07-19 15:16
    '����:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If frmInTend Is Nothing Then Exit Sub
    Call frmInTend.SetFontSize(bytSize)
End Sub