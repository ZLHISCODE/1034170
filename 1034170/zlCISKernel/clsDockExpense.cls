VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Activate() '�Ӵ��弤��
Public Event RequestRefresh() 'Ҫ��������ˢ��
Public Event StatusTextUpdate(ByVal Text As String) 'Ҫ�����������״̬������
Private WithEvents mfrmExpense As frmDockExpense
Attribute mfrmExpense.VB_VarHelpID = -1
Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:���˺�
    '����:2012-06-18 16:50:35
    '����:50793
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmExpense Is Nothing Then Exit Sub
    mfrmExpense.SetFontSize (bytSize)
End Sub
Private Sub Class_Initialize()
    Set mfrmExpense = New frmDockExpense
End Sub

Private Sub Class_Terminate()
    Unload mfrmExpense
    Set mfrmExpense = Nothing
    Set grsҽ�Ƹ��ʽ = Nothing
End Sub

Private Sub mfrmExpense_Activate()
    RaiseEvent Activate
End Sub

Private Sub mfrmExpense_RequestRefresh()
    RaiseEvent RequestRefresh
End Sub

Private Sub mfrmExpense_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Public Function zlGetForm() As Object
'���ܣ���ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ

    '��ʼ��ҽ������:�ɷ�ɢ�������GetForm��ȥ��ʼ��
    Call gclsInsure.InitOracle(gcnOracle)

    '��ʼ��ȡȨ��,�Լӿ�����ٶ�(���Ǳ�Ҫ��)
    Call GetInsidePrivs(pҽ�����ѹ���, True)
    
    Set zlGetForm = mfrmExpense
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, Optional ByRef objSquareCard As Object)
'���ܣ������Ӵ���Ĳ˵��͹�����(����������Ҫʹ�õĲ˵��͹�����)
'������CommandBars=�����ڲ鿴ʱ���Բ���(����Nothing)
'      objSquareCard=ͨ��ҽ������վ�����һ��ͨ���㲿��
    Call mfrmExpense.zlDefCommandBars(frmParent, CommandBars, objSquareCard)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mfrmExpense.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mfrmExpense.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mfrmExpense.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal lng����id As Long, ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, Optional ByVal blnMoved As Boolean, Optional ByVal bln����ִ�� As Boolean)
'���ܣ�ˢ��ҽ����������
'������lng����ID=��ǰҽ������ID
'      lngҽ��ID=�൱��ҽ��վ��ǰ��ID
'      blnMoved=�ò��˵������Ƿ���ת��
'      bln����ִ��=���ڼ�����Ŀ��һ���ɼ���һ����Ŀ���Ƿ�������е�ĳһ������ִ��
    Call mfrmExpense.zlRefresh(lng����id, lngҽ��ID, lng���ͺ�, blnMoved, bln����ִ��)
End Sub