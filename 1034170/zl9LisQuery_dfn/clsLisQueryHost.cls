VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisQueryHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get OwnerFormHandle() As Long
    '������ľ��
End Property

Public Sub RaiseFinished(objQuery As clsLisQuery)
    '�ڿͻ��˴������¼�����֪ͨ�����򣬳��������
End Sub

Public Function GetRecordSet(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    '���Ÿ��ͻ��ˣ�����ִ��ȡ����
End Function

Public Function ClientTrigger(ByVal Index As Long, ByVal strAction As String, strData As String) As String
    '���Ÿ��ͻ��˴����¼�
    '           Ԥ������
End Function

