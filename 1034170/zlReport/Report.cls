VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarϵͳ As Long '�ֲ�����

Private mvar��� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar˵�� As String '�ֲ�����
Private mvar��ӡ�� As String '�ֲ�����
Private mvar��ֽ As Integer   '�ֲ�����
Private mvarƱ�� As Boolean '�ֲ�����
Private mvar��ӡ��ʽ As Integer '�ֲ�����
Private mvar�޸�ʱ�� As Date '�ֲ�����
Private mvarItems As RPTItems '�ֲ�����
Private mvarDatas As RPTDatas '�ֲ�����
Private mvarFmts As RPTFmts '�ֲ�����
Private mvar��ֹ��ʼʱ�� As Date '�ֲ�����
Private mvar��ֹ����ʱ�� As Date '�ֲ�����

Private mvarbytFormat As Byte '�ֲ�����
Private mvarblnLoad As Boolean  '�ֲ�����
Private mvarintGridCount As Integer
Private mvarintGridID As Integer

Public Property Let �޸�ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�޸�ʱ�� = 5
    mvar�޸�ʱ�� = vData
End Property

Public Property Get �޸�ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�޸�ʱ��
    �޸�ʱ�� = mvar�޸�ʱ��
End Property

Public Property Let ϵͳ(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.intGridID = 5
    mvarϵͳ = vData
End Property

Public Property Get ϵͳ() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.intGridID
    ϵͳ = mvarϵͳ
End Property

Public Property Let intGridID(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.intGridID = 5
    mvarintGridID = vData
End Property

Public Property Get intGridID() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.intGridID
    intGridID = mvarintGridID
End Property

Public Property Let intGridCount(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.intGridCount = 5
    mvarintGridCount = vData
End Property

Public Property Get intGridCount() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.intGridCount
    intGridCount = mvarintGridCount
End Property

Public Property Let bytFormat(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.bytFormat = 5
    mvarbytFormat = vData
End Property

Public Property Get bytFormat() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.bytFormat
    bytFormat = mvarbytFormat
End Property

Public Property Let blnLoad(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.blnLoad = 5
    mvarblnLoad = vData
End Property

Public Property Get blnLoad() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.blnLoad
    blnLoad = mvarblnLoad
End Property
'-----------------------------------------------------

Public Property Set Fmts(ByVal vData As RPTFmts)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Fmts = Form1
    Set mvarFmts = vData
End Property

Public Property Get Fmts() As RPTFmts
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Fmts
    Set Fmts = mvarFmts
End Property

Public Property Set Datas(ByVal vData As RPTDatas)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Datas = Form1
    Set mvarDatas = vData
End Property

Public Property Get Datas() As RPTDatas
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Datas
    Set Datas = mvarDatas
End Property

Public Property Set Items(ByVal vData As RPTItems)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Items = Form1
    Set mvarItems = vData
End Property

Public Property Get Items() As RPTItems
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Items
    On Error Resume Next
    Set Items = mvarItems
    On Error GoTo 0
End Property

Public Property Let ��ӡ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ӡ�� = 5
    mvar��ӡ�� = vData
End Property

Public Property Get ��ӡ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ӡ��
    ��ӡ�� = mvar��ӡ��
End Property

Public Property Let ��ֽ(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ֽ = 5
    mvar��ֽ = vData
End Property

Public Property Get ��ֽ() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ֽ
    ��ֽ = mvar��ֽ
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

Public Property Let ���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��� = vData
End Property

Public Property Get ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���
    ��� = mvar���
End Property

Private Sub Class_Initialize()
    '�Զ���ʼ��
    Set Datas = New RPTDatas
    Set Items = New RPTItems
    Set Fmts = New RPTFmts
End Sub

Public Property Let Ʊ��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ʊ�� = 5
    mvarƱ�� = vData
End Property

Public Property Get Ʊ��() As Boolean
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

Public Property Let ��ֹ��ʼʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�޸�ʱ�� = 5
    mvar��ֹ��ʼʱ�� = vData
End Property

Public Property Get ��ֹ��ʼʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�޸�ʱ��
    ��ֹ��ʼʱ�� = mvar��ֹ��ʼʱ��
End Property

Public Property Let ��ֹ����ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�޸�ʱ�� = 5
    mvar��ֹ����ʱ�� = vData
End Property

Public Property Get ��ֹ����ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�޸�ʱ��
    ��ֹ����ʱ�� = mvar��ֹ����ʱ��
End Property