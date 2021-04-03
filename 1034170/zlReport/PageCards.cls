VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'�ֲ����������漯��
Private mCol As Collection

Public Function Add(id As Integer, X As Long, Y As Long, W As Long, H As Long, RowB As Long, RowE As Long, Row As Long, Col As Long, Item As Collection, Optional sKey As String) As PageCard
    '�����¶���
    Dim objNewMember As PageCard
    Set objNewMember = New PageCard

    '���ô��뷽��������
    objNewMember.Key = Mid(sKey, 2) 'KeyΪ������ؼ���,Ӧ��ID��Ӧ
    objNewMember.id = id '���ݱ���ؼ�����
    
    '�ñ����ڸ�ҳ�����������λ�á��ߴ�(����ͷ)
    objNewMember.X = X
    objNewMember.Y = Y
    objNewMember.W = W
    objNewMember.H = H '����������ø߶�(����ͷ),���=FixH,��ʾ�������ͷ
  
    objNewMember.RowB = RowB '�����ڸ�ҳ�Ŀ�ʼ��(�ǹ̶���)
    objNewMember.RowE = RowE '�����ڸ�ҳ�Ľ�����(�ǹ̶���)
    objNewMember.Row = Row
    objNewMember.Col = Col
    objNewMember.Item = Item
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As PageCard
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
     On Error Resume Next
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Items() As Collection
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Items = mCol
End Property

Public Property Get count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    count = mCol.count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    '������������ For...Each �﷨ö�ٸü��ϡ�
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    '������󴴽�����
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    '����ֹ���ƻ�����
    Set mCol = Nothing
End Sub
