VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'��������ֵ�ľֲ�����
Private mvar���� As Integer '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar���� As Boolean '�ֲ�����
Private mvar�ӿڳ����� As String '�ֲ�����
'��������ֵ�ľֲ�����
Private mvar���� As Integer '�ֲ�����
Private mvar���� As Integer
Private mvar�Զ���ȡ As Integer

Public Property Let �Ƿ��Զ���ȡ(ByVal vdata As Integer)
    mvar�Զ���ȡ = vdata
End Property

Public Property Get �Ƿ��Զ���ȡ() As Integer
    �Ƿ��Զ���ȡ = mvar�Զ���ȡ
End Property

Public Property Let ����(ByVal vdata As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vdata
End Property

Public Property Get ����() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let �ɷ�����(ByVal vdata As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vdata
End Property


Public Property Get �ɷ�����() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    �ɷ����� = mvar����
End Property



Public Property Let �ӿڳ�����(ByVal vdata As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ӿڳ����� = 5
    mvar�ӿڳ����� = vdata
End Property


Public Property Get �ӿڳ�����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ӿڳ�����
    �ӿڳ����� = mvar�ӿڳ�����
End Property



Public Property Let ����(ByVal vdata As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vdata
End Property


Public Property Get ����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vdata As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vdata
End Property


Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vdata As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vdata
End Property

Public Property Get ����() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property


