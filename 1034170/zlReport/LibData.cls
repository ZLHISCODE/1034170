VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'��������ֵ�ľֲ�����
Private mvarDataSet As ADODB.Recordset '�ֲ�����
Private mvarDataName As String '�ֲ�����

Public Property Let DataName(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.DataName = 5
    mvarDataName = vData
End Property

Public Property Get DataName() As String
Attribute DataName.VB_UserMemId = 0
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.DataName
    DataName = mvarDataName
End Property

Public Property Set DataSet(ByVal vData As ADODB.Recordset)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.DataSet = Form1
    Set mvarDataSet = vData
End Property

Public Property Get DataSet() As ADODB.Recordset
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.DataSet
    Set DataSet = mvarDataSet
End Property


