VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageCells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"PageCell"
Attribute VB_Ext_KEY = "Member0" ,"PageCell"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'�ֲ����������漯��
Private mCol As Collection

Public Function Add(ID As Integer, X As Long, Y As Long, W As Long, H As Long, MaxH As Long, RowB As Long, RowE As Long, VRowE As Long, ColB As Long, ColE As Long, FixW As Long, FixH As Long, Copys As Byte, Optional sKey As String) As PageCell
    '�����¶���
    Dim objNewMember As PageCell
    Set objNewMember = New PageCell

    '���ô��뷽��������
    objNewMember.Key = Mid(sKey, 2) 'KeyΪ������ؼ���,Ӧ��ID��Ӧ
    objNewMember.ID = ID '���ݱ���ؼ�����
    
    '�ñ����ڸ�ҳ�����������λ�á��ߴ�(����ͷ)
    objNewMember.X = X
    objNewMember.Y = Y
    objNewMember.W = W
    objNewMember.H = H '����������ø߶�(����ͷ),���=FixH,��ʾ�������ͷ
    
    '�ñ������Ƹ߶�(�����ӱ���)
    objNewMember.MaxH = MaxH
    
    objNewMember.FixW = FixW '����̶��п���(���Ϊ0,��ʾ��ҳû�л򲻴�ӡ�̶���)
    objNewMember.FixH = FixH '����̶��и߶�(���Ϊ0,��ʾ��ҳû�л򲻴�ӡ�̶���)
    objNewMember.RowB = RowB '�����ڸ�ҳ�Ŀ�ʼ��(�ǹ̶���)
    objNewMember.RowE = RowE '�����ڸ�ҳ�Ľ�����(�ǹ̶���)
    objNewMember.VRowE = VRowE '�����ڸ�ҳ�����������(0,����>RowE,�����ǳ����˱����з�Χ��)�����ڷ�������Ʊ��ʱ����
    objNewMember.ColB = ColB '�����ڸ�ҳ�Ŀ�ʼ��(�ǹ̶���)
    objNewMember.ColE = ColE '�����ڸ�ҳ�Ľ�����(�ǹ̶���)
    objNewMember.Copys = Copys '����ķ�����,��ҳ������ֹ��Χ�Ѿ������˷���������
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As PageCell
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
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