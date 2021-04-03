VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BillDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"BillDetail"
Attribute VB_Ext_KEY = "Member0" ,"BillDetail"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 'Ҫ���������

'�ֲ����������漯��
Private mCol As Collection

Public Function Add(InComes As BillInComes, Detail As Detail, �շ�ϸĿID As Long, _
    ��� As Integer, �������� As Integer, �շ���� As String, ���㵥λ As String, _
    ���� As Integer, ���� As Double, ���ӱ�־ As Byte, ִ�в���ID As Long, ��ҩ���� As String, _
    Optional ������Ŀ�� As Boolean, Optional ���մ���ID As Long, Optional ���ձ��� As String, _
    Optional ժҪ As String, Optional sKey As String) As BillDetail
    
    '�����¶���
    Dim objNewMember As BillDetail
    Set objNewMember = New BillDetail

    '���ô��뷽��������
    objNewMember.Key = sKey
    
    Set objNewMember.InComes = InComes
    Set objNewMember.Detail = Detail
    
    objNewMember.�շ�ϸĿID = �շ�ϸĿID
    objNewMember.��� = ���
    objNewMember.�������� = ��������
    objNewMember.�շ���� = �շ����
    objNewMember.���㵥λ = ���㵥λ
    objNewMember.���� = ����
    objNewMember.���� = ����
    objNewMember.���ӱ�־ = ���ӱ�־
    objNewMember.ִ�в���ID = ִ�в���ID
    objNewMember.��ҩ���� = ��ҩ����
    objNewMember.������Ŀ�� = ������Ŀ��
    objNewMember.���մ���ID = ���մ���ID
    objNewMember.���ձ��� = ���ձ���
    objNewMember.ժҪ = ժҪ
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As BillDetail
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