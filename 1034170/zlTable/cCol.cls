VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'################################################################################################################
'## �ڲ�����
'################################################################################################################

Private mvarKey As String               '�ؼ���
Private mvarCol As Long                 '��
Private mvarStartX As Long              'X����
Private mvarWidth As Long               '�п�
Private mvarFixedWidth As Boolean       '�Ƿ��������ȣ�Ĭ��ΪFalse
Private mvarMaxWidth As Long            '����п�
Private mvarVisibled As Boolean         '�Ƿ�ɼ�
Private mvarTitle As String             '�б���
Private mvarTag As String               '��չ���
Private mvarHAlignment As HAlignEnum    'ˮƽ���뷽ʽ
Private mvarVAlignment As VAlignEnum    '��ֱ���뷽ʽ
Private mvarTextFormat As String        '�ı���ʽ
Private mvarIndent As Long              '����ֵ
Private mvarProtected As Boolean        '�Ƿ񱣻�

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let Col(ByVal vData As Long)
    mvarCol = vData
End Property

Public Property Get Col() As Long
    Col = mvarCol
End Property

Public Property Let StartX(ByVal vData As Long)
    mvarStartX = vData
End Property

Public Property Get StartX() As Long
    StartX = mvarStartX
End Property

Public Property Let Width(ByVal vData As Long)
    mvarWidth = vData
End Property

Public Property Get Width() As Long
    Width = mvarWidth
End Property

Public Property Let FixedWidth(ByVal vData As Boolean)
    mvarFixedWidth = vData
End Property

Public Property Get FixedWidth() As Boolean
    FixedWidth = mvarFixedWidth
End Property

Public Property Let MaxWidth(ByVal vData As Long)
    mvarMaxWidth = vData
End Property

Public Property Get MaxWidth() As Long
    MaxWidth = mvarMaxWidth
End Property

Public Property Let Visibled(ByVal vData As Boolean)
    mvarVisibled = vData
End Property

Public Property Get Visibled() As Boolean
    Visibled = mvarVisibled
End Property

Public Property Let Title(ByVal vData As String)
    mvarTitle = vData
End Property

Public Property Get Title() As String
    Title = mvarTitle
End Property

Public Property Let Tag(ByVal vData As String)
    mvarTag = vData
End Property

Public Property Get Tag() As String
    Tag = mvarTag
End Property

Public Property Let HAlignment(ByVal vData As HAlignEnum)
    mvarHAlignment = vData
End Property

Public Property Get HAlignment() As HAlignEnum
    HAlignment = mvarHAlignment
End Property

Public Property Let VAlignment(ByVal vData As VAlignEnum)
    mvarVAlignment = vData
End Property

Public Property Get VAlignment() As VAlignEnum
    VAlignment = mvarVAlignment
End Property

Public Property Let TextFormat(ByVal vData As String)
    mvarTextFormat = vData
End Property

Public Property Get TextFormat() As String
    TextFormat = mvarTextFormat
End Property

Public Property Let Indent(ByVal vData As Long)
    mvarIndent = vData
End Property

Public Property Get Indent() As Long
    Indent = mvarIndent
End Property

Public Property Let Protected(ByVal vData As Boolean)
    mvarProtected = vData
End Property

Public Property Get Protected() As Boolean
    Protected = mvarProtected
End Property

'################################################################################################################
'## ���ܣ�  ���ƶ���
'##
'## ������  blnWithKey  :�Ƿ���ͬKeyֵһ���ƣ�
'################################################################################################################
Public Function Clone(Optional blnWithKey As Boolean = False) As cCol
    Dim cNode As New cCol
    If blnWithKey Then cNode.Key = Me.Key
    cNode.Col = Me.Col
    cNode.StartX = Me.StartX
    cNode.Width = Me.Width
    cNode.FixedWidth = Me.FixedWidth
    cNode.MaxWidth = Me.MaxWidth
    cNode.Visibled = Me.Visibled
    cNode.Title = Me.Title
    cNode.Tag = Me.Tag
    cNode.HAlignment = Me.HAlignment
    cNode.VAlignment = Me.VAlignment
    cNode.TextFormat = Me.TextFormat
    cNode.Indent = Me.Indent
    cNode.Protected = Me.Protected
    
    Set Clone = cNode
End Function

Private Sub Class_Initialize()
    'Ĭ�����Գ�ʼ��
    Me.Width = 1600             'Ĭ���п�1600
    Me.FixedWidth = False       'Ĭ�Ͽ��Ե����п�
    Me.MaxWidth = 999999        'Ĭ�������Ȳ�����999999
    Me.Visibled = True          'Ĭ����ʾ
    Me.Title = ""               'Ĭ���б���Ϊ��
    Me.Tag = ""                 'Ĭ���б��Ϊ��
    Me.HAlignment = HALignLeft  'Ĭ�Ͽ���
    Me.VAlignment = VALignTop   'Ĭ�Ͽ���
    Me.TextFormat = ""          'Ĭ�ϸ�ʽΪ��
    Me.Indent = 0               'Ĭ������Ϊ0
    Me.Protected = False        'Ĭ�ϲ�����
End Sub
















