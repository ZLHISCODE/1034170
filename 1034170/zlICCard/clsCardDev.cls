VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCardDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'��������ֵ�ľֲ�����

'�ӿ���
Public Sub SetCard(Optional strInput As String, Optional strOut As String)
    '����ģ�壬��ʵ�־��幦��
End Sub

Public Function ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    '����ģ�壬��ʵ�־��幦��
End Function

Public Function initCard(Optional strInput As String, Optional strOut As String) As Boolean
    '����ģ�壬��ʵ�־��幦��
End Function

Public Function CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    '����ģ�壬��ʵ�־��幦��
End Function

Public Function GetSpare() As Currency
    '����ģ�壬����һ��ͨ�Ŀ���ǰ���
End Function

Public Function PaymentSwap(ByVal dblMoney As Double, ByRef dblBalance As Double, ByRef intCardType As Integer, _
    ByVal lngHOSCode As Long, ByRef strCardNO As String, ByRef strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean
    '����ģ�壬ִ��һ��ͨ�۷ѽ���
End Function

Public Function ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean
    '����ģ�壬ִ��һ��ͨ�˷ѽ���
End Function

Public Function UploadSwap(ByVal lng����ID As Long, ByRef strOut As String) As Boolean
    '����ģ�壬��������������ϴ��Ĺ�����Ŀǰ��Ӧ��������һ��ͨ����ɲ�����Ϣ������Ϣ���ϴ�
    
End Function

Private Sub Class_Terminate()
    Call CloseCard
End Sub