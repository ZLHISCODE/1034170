VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisQuery_Base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements zl9LisQuery_Def.clsLisQuery
Private objHost   As zl9LisQuery_Def.clsLisQueryHost
Private udeResult As String
Private varValue  As Variant
Private mlngIndex As Long
Private mfrmMain As frmMain

Private Sub Class_Initialize()
    If mfrmMain Is Nothing Then Set mfrmMain = New frmMain
End Sub

Private Sub Class_Terminate()
    If Not mfrmMain Is Nothing Then
        Unload mfrmMain
        Set mfrmMain = Nothing
    End If
End Sub

Private Property Let clsLisQuery_Index(ByVal RHS As Long)
    mlngIndex = RHS
End Property

Private Property Get clsLisQuery_Index() As Long
    clsLisQuery_Index = mlngIndex
End Property

Private Property Get clsLisQuery_Name() As String
    '�˵�����
    clsLisQuery_Name = "�ۺϲ�ѯ(&Q)"
End Property

Private Sub clsLisQuery_DoAction(ByVal conShowMode As zl9LisQuery_Def.QueryShowMode)
    'ʵ�ֲ�ѯ���ܵĴ���
    If mfrmMain Is Nothing Then Set mfrmMain = New frmMain
    Call mfrmMain.ShowMe(mlngIndex, conShowMode, objHost)
   
End Sub

Private Property Get clsLisQuery_Description() As String
    '�����˵����Ϣ
    clsLisQuery_Description = "ZLLIS�ۺϲ�ѯ���"
End Property

Private Sub clsLisQuery_InitQuery(QueryHost As zl9LisQuery_Def.clsLisQueryHost)
    '��ʼ��
    Set objHost = QueryHost
End Sub

Private Property Get clsLisQuery_Result() As String
    '����
    clsLisQuery_Result = udeResult
End Property

Private Property Get clsLisQuery_Value() As Variant
    '����ֵ
    clsLisQuery_Value = varValue
End Property
