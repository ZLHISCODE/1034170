VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRow"
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
Private mvarRow As Long                 '��
Private mvarStartY As Long              'Y����
Private mvarHeight As Long              '�и�
Private mvarMultiLine As Boolean        '�Ƿ������ʾ��Ĭ��ΪTrue
Private mvarAutoHeight As Boolean       '�Զ��߶ȣ�Ĭ��ΪTrue
Private mvarMaxHeight As Long           '����и�
Private mvarVisibled As Boolean         '�Ƿ�ɼ�
Private mvarTitle As String             '�б���
Private mvarTag As String               '��չ���
Private mvarHAlignment As HAlignEnum    'ˮƽ���뷽ʽ
Private mvarVAlignment As VAlignEnum    '��ֱ���뷽ʽ
Private mvarTextFormat As String        '�ı���ʽ
Private mvarProtected As Boolean        '�Ƿ񱣻�

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let Row(ByVal vData As Long)
    mvarRow = vData
End Property

Public Property Get Row() As Long
    Row = mvarRow
End Property

Public Property Let StartY(ByVal vData As Long)
    mvarStartY = vData
End Property

Public Property Get StartY() As Long
    StartY = mvarStartY
End Property

Public Property Let Height(ByVal vData As Long)
    mvarHeight = vData
End Property

Public Property Get Height() As Long
    Height = mvarHeight
End Property

Public Property Let MultiLine(ByVal vData As Boolean)
    mvarMultiLine = vData
End Property

Public Property Get MultiLine() As Boolean
    MultiLine = mvarMultiLine
End Property

Public Property Let AutoHeight(ByVal vData As Boolean)
    mvarAutoHeight = vData
End Property

Public Property Get AutoHeight() As Boolean
    AutoHeight = mvarAutoHeight
End Property

Public Property Let MaxHeight(ByVal vData As Long)
    mvarMaxHeight = vData
End Property

Public Property Get MaxHeight() As Long
    MaxHeight = mvarMaxHeight
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
Public Function Clone(Optional blnWithKey As Boolean = False) As cRow
    Dim cNode As New cRow
    If blnWithKey Then cNode.Key = Me.Key
    cNode.Row = Me.Row
    cNode.StartY = Me.StartY
    cNode.Height = Me.Height
    cNode.MultiLine = Me.MultiLine
    cNode.AutoHeight = Me.AutoHeight
    cNode.MaxHeight = Me.MaxHeight
    cNode.Visibled = Me.Visibled
    cNode.Title = Me.Title
    cNode.Tag = Me.Tag
    cNode.HAlignment = Me.HAlignment
    cNode.VAlignment = Me.VAlignment
    cNode.TextFormat = Me.TextFormat
    cNode.Protected = Me.Protected
    
    Set Clone = cNode
End Function

Private Sub Class_Initialize()
    'Ĭ�����Գ�ʼ��
    Me.Height = 300             'Ĭ���и�300
    Me.MultiLine = True         'Ĭ�϶�����ʾ
    Me.AutoHeight = True        'Ĭ���Զ��и�
    Me.MaxHeight = 999999       'Ĭ������и�999999
    Me.Visibled = True          'Ĭ����ʾ
    Me.Title = ""               'Ĭ���б���Ϊ��
    Me.Tag = ""                 'Ĭ���б��Ϊ��
    Me.HAlignment = HALignLeft  'Ĭ�Ͽ���
    Me.VAlignment = VALignTop   'Ĭ�Ͽ���
    Me.TextFormat = ""          'Ĭ�ϸ�ʽΪ��
    Me.Protected = False        'Ĭ�ϲ�����
End Sub


















