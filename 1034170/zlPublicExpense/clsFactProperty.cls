VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFactProperty"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'��������ֵ�ľֲ�����
Private mvar��������ID As Long '�ֲ�����
Private mvarʹ����� As String '�ֲ�����
Private mvar��ӡ��ʽ As Integer '�ֲ�����
Private mvar��ӡ��ʽ As Integer '�ֲ�����
'��������ֵ�ľֲ�����
Private mvarƱ�� As Integer '�ֲ�����
Private mvar�ϸ���� As Boolean '�ֲ�����
Private mvar����ʹ����� As Boolean '�ֲ�����
Private mvarƱ�ų��� As Integer   '�ֲ�����
Private mvarLastUseID As Long     '�ֲ�����

Public Property Let LastUseID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.LastUseID = 5
    mvarLastUseID = vData
End Property
Public Property Get LastUseID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.LastUseID
    LastUseID = mvarLastUseID
End Property


Public Property Let Ʊ�ų���(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ʊ�ų��� = 5
    mvarƱ�ų��� = vData
End Property
Public Property Get Ʊ�ų���() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ʊ�ų���
    Ʊ�ų��� = mvarƱ�ų���
End Property


Public Property Let ����ʹ�����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ʹ����� = 5
    mvar����ʹ����� = vData
End Property


Public Property Get ����ʹ�����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ʹ�����
    ����ʹ����� = mvar����ʹ�����
End Property


Public Property Let �ϸ����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ϸ���� = 5
    mvar�ϸ���� = vData
End Property


Public Property Get �ϸ����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ϸ����
    �ϸ���� = mvar�ϸ����
End Property



Public Property Let Ʊ��(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ʊ�� = 5
    mvarƱ�� = vData
End Property


Public Property Get Ʊ��() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ʊ��
      Ʊ�� = mvarƱ��
End Property



Public Property Let ��ӡ��ʽ(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ӡ��ʽ = 5
    mvar��ӡ��ʽ = vData
End Property


Public Property Get ��ӡ��ʽ() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ӡ��ʽ
      ��ӡ��ʽ = mvar��ӡ��ʽ
End Property



Public Property Let ��ӡ��ʽ(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ӡ��ʽ = 5
    mvar��ӡ��ʽ = vData
End Property


Public Property Get ��ӡ��ʽ() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ӡ��ʽ
    ��ӡ��ʽ = mvar��ӡ��ʽ
End Property



Public Property Let ʹ�����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ʹ����� = 5
    mvarʹ����� = vData
End Property


Public Property Get ʹ�����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ʹ�����
    ʹ����� = mvarʹ�����
End Property



Public Property Let ��������ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��������ID = 5
    mvar��������ID = vData
End Property


Public Property Get ��������ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������ID
    ��������ID = mvar��������ID
End Property


