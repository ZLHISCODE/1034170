VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�̳д����¼�
Public Event Activate()
Public Event AfterSaved(ByVal lngOrderId As Long, ByVal lngSaveType As Long)
Public Event AfterOpen(ByVal intEditType As EditTypeEnum)
Public Event AfterClosed(ByVal lngOrderId As Long)
Public Event AfterPrinted(ByVal lngOrderId As Long)
Public Event AfterDeleted(ByVal lngOrderId As Long)

Private WithEvents frmReport As frmDockReport
Attribute frmReport.VB_VarHelpID = -1

Private Sub Class_initialize()
    Set frmReport = New frmDockReport
End Sub

Private Sub Class_Terminate()
    Unload frmReport
    Set frmReport = Nothing
End Sub

Private Sub frmReport_Activate()
    RaiseEvent Activate
End Sub
Private Sub frmReport_AfterOpen(ByVal intEditType As EditTypeEnum)
    RaiseEvent AfterOpen(intEditType)
End Sub

Private Sub frmReport_AfterSaved(ByVal lngOrderId As Long, ByVal lngSaveType As Long)
     RaiseEvent AfterSaved(lngOrderId, lngSaveType)
End Sub
Private Sub frmReport_AfterClosed(ByVal lngOrderId As Long)
    RaiseEvent AfterClosed(lngOrderId)
End Sub
Private Sub frmReport_AfterPrinted(ByVal lngOrderId As Long)
    RaiseEvent AfterPrinted(lngOrderId)
End Sub
Private Sub frmReport_AfterDeleted(ByVal lngOrderId As Long)
    RaiseEvent AfterDeleted(lngOrderId)
End Sub
Public Function zlGetForm() As Object
    Set zlGetForm = frmReport
End Function

Public Sub zlDefCommandBars(ByVal CommandBars As Object)
    Call frmReport.zlDefCommandBars(CommandBars)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmReport.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmReport.zlExecuteCommandBars(Control)
End Sub

Public Sub ConfigPrint(ByVal strPrintDevice As String, ByVal lngCopies As Long)
    Call frmReport.ConfigPrint(strPrintDevice, lngCopies)
End Sub

Public Function zlRefresh(ByVal lngOrderId As Long, ByVal lngDeptId As Long, Optional ByVal bnEdit As Boolean, _
                        Optional ByVal blnMoved As Boolean, Optional ByVal blnCanPrint As Boolean = True, Optional ByVal lngModule As Long, Optional ByVal blnForce As Boolean) As Long
    '����:����ˢ��ָ��ҽ��id�ı������ݣ�����������ṩ�༭����
    '����:  lngOrderId-ҽ��id;
    '       lngDeptId-��ǰ�������ţ�ע�ⲻ�ǲ��˱��ξ�����ң�
    '       blnEdit-�Ƿ������༭�����ߵ���ʷ����Ӧ����Ϊ��
    '       blnMoved-�����Ƿ�ת��
    '       �쳣����0,���򷵻�1
    zlRefresh = frmReport.zlRefresh(lngOrderId, lngDeptId, bnEdit, blnForce, blnMoved, blnCanPrint, lngModule)
End Function

Public Sub RefPacsPic()
    '����:ˢ�µ�ǰ���ڱ༭�����PACSͼƬ
    Call frmReport.RefPacsPic
End Sub
Public Property Let PacsCore(objCore As Object)
    Set gobjPacsCore = objCore
End Property
Public Property Set PacsCore(objCore As Object)
    Set gobjPacsCore = objCore
End Property