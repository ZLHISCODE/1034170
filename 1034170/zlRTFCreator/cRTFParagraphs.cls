VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFParagraphs"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cRTFParagraph"
Attribute VB_Ext_KEY = "Member0" ,"cRTFParagraph"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection

Public Function GetTextRTF() As String
    Dim i As Long, strR As String
    For i = 1 To mCol.Count
        strR = strR & mCol(i).GetTextRTF
    Next
    GetTextRTF = strR
End Function

Public Function Add(Optional sKey As String, Optional IsInTable As Boolean = False, _
    Optional SpaceBefore As Long = 0, Optional SpaceAfter As Long = 0, _
    Optional SpaceLine As Long = 0, Optional SpaceFirstLine As Long = 0, _
    Optional FontSizeDefault As Long = 0) As Long
    '�����¶���
    Dim objNewMember As cRTFParagraph
    Set objNewMember = New cRTFParagraph


    '���ô��뷽��������
    objNewMember.IsInTable = IsInTable
    objNewMember.SpaceBefore = SpaceBefore
    objNewMember.SpaceAfter = SpaceAfter
    objNewMember.SpaceLine = SpaceLine
    objNewMember.SpaceFirstLine = SpaceFirstLine
    objNewMember.FontSizeDefault = FontSizeDefault
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '�����Ѵ����Ķ���
    Add = mCol.Count
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As cRTFParagraph
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
