VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Private mvar���� As String '�ֲ�����
Private mvarSQL As String '�ֲ�����
Private mvar�ֶ� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar���� As Byte '�ֲ�����
Private mvar˵�� As String '�ֲ�����
Private mvarPars As RPTPars '�ֲ�����
Private mvarԭ���� As String '�ֲ�����
Private mintDataConnNo As Integer         '�������ӱ��

Public Property Set Pars(ByVal vData As RPTPars)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Pars = Form1
    Set mvarPars = vData
End Property

Public Property Get Pars() As RPTPars
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Pars
    Set Pars = mvarPars
End Property

Public Property Let ����(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let ˵��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.˵�� = 5
    mvar˵�� = vData
End Property

Public Property Get ˵��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.˵��
    ˵�� = mvar˵��
End Property

Public Property Let SQL(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SQL = 5
    mvarSQL = vData
End Property

Public Property Get SQL() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SQL
    SQL = mvarSQL
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Private Sub Class_Initialize()
    Set Pars = New RPTPars
End Sub

Public Property Let �ֶ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ֶ� = 5
    mvar�ֶ� = vData
End Property

Public Property Get �ֶ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ֶ�
    �ֶ� = mvar�ֶ�
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ֶ� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ֶ�
    ���� = mvar����
End Property

Public Property Let ԭ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvarԭ���� = vData
End Property

Public Property Get ԭ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ԭ���� = mvarԭ����
End Property

Public Property Let �������ӱ��(ByVal intDataConnNo As Integer)
    mintDataConnNo = intDataConnNo
End Property

Public Property Get �������ӱ��() As Integer
    �������ӱ�� = mintDataConnNo
End Property
