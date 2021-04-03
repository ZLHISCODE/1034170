VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExpenseBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ExpenseBill"
Attribute VB_Ext_KEY = "Member0" ,"ExpenseBill"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 'Ҫ���������

'�ֲ����������漯��
Private mCol As Collection

Public Function Add(NO As String, ����ID As Long, ��ҳID As Long, ����ID As Long, ����ID As Long, ���� As String, �Ա� As String, ���� As String, ��ʶ�� As String, ���� As String, �ѱ� As String, �����־ As Byte, �Ӱ��־ As Byte, Ӥ���� As Byte, �ಡ�˵� As Boolean, ��������ID As Long, ������ As String, ������ As String, ����Ա���� As String, ����Ա��� As String, ����ʱ�� As Date, �Ǽ�ʱ�� As Date, Details As BillDetails, Optional sKey As String) As ExpenseBill
    '�����¶���
    Dim objNewMember As ExpenseBill
    Set objNewMember = New ExpenseBill

    '���ô��뷽��������
    objNewMember.Key = sKey
    objNewMember.NO = NO
    objNewMember.����ID = ����ID
    objNewMember.��ҳID = ��ҳID
    objNewMember.����ID = ����ID
    objNewMember.����ID = ����ID
    objNewMember.���� = ����
    objNewMember.�Ա� = �Ա�
    objNewMember.���� = ����
    objNewMember.��ʶ�� = ��ʶ��
    objNewMember.���� = ����
    objNewMember.�ѱ� = �ѱ�
    objNewMember.�����־ = �����־
    objNewMember.�Ӱ��־ = �Ӱ��־
    objNewMember.Ӥ���� = Ӥ����
    objNewMember.�ಡ�˵� = �ಡ�˵�
    objNewMember.��������ID = ��������ID
    objNewMember.������ = ������
    objNewMember.������ = ������
    objNewMember.����Ա���� = ����Ա����
    objNewMember.����Ա��� = ����Ա���
    objNewMember.����ʱ�� = ����ʱ��
    objNewMember.�Ǽ�ʱ�� = �Ǽ�ʱ��
    Set objNewMember.Details = Details
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As ExpenseBill
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��'debug.print x.Count
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