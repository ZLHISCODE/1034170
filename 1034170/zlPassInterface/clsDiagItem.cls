VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiagItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstr�������� As String
Private mstr��ϱ��� As String
Private mstr������� As String
Private mstr���ID As String
Private mstr����ID As String
Private mstr������� As String
Private mstr���ʱ�� As String

Public Property Let str��������(ByVal strData As String)
    mstr�������� = strData
End Property

Public Property Get str��������() As String
    str�������� = mstr��������
End Property

Public Property Let str��ϱ���(ByVal strData As String)
    mstr��ϱ��� = strData
End Property

Public Property Get str��ϱ���() As String
    str��ϱ��� = mstr��ϱ���
End Property

Public Property Let str�������(ByVal strData As String)
    mstr������� = strData
End Property

Public Property Get str�������() As String
    str������� = mstr�������
End Property
Public Property Let str����ID(ByVal strData As String)
    mstr����ID = strData
End Property
Public Property Get str����ID() As String
    str����ID = mstr����ID
End Property
Public Property Let str���ID(ByVal strData As String)
    mstr���ID = strData
End Property
Public Property Get str���ID() As String
    str���ID = mstr���ID
End Property
Public Property Let str�������(ByVal strData As String)
    mstr������� = strData
End Property
Public Property Get str�������() As String
    str������� = mstr�������
End Property
Public Property Let str���ʱ��(ByVal strData As String)
    mstr���ʱ�� = strData
End Property
Public Property Get str���ʱ��() As String
    str���ʱ�� = mstr���ʱ��
End Property