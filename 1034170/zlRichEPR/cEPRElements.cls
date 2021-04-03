VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cEPRElement"
Attribute VB_Ext_KEY = "Member0" ,"cEPRElement"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mlngKey As Long

Public Property Get Item(vntIndexKey As Variant) As cEPRElement
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

'################################################################################################################
'## ���ܣ�  ������ж����ID�Ƿ���Ч
'##
'## ������  eEditType :��ǰ�ı༭ģʽ
'##
'## ˵����  ���û��ID�������ݿ�����ȡһ��ΨһID�š�
'################################################################################################################
Public Sub CheckValidIDs(eEditType As EditTypeEnum)
    Dim i As Long
    Select Case eEditType
    Case cprET_�����ļ�����
        For i = 1 To mCol.Count
            If mCol(i).ID = 0 Then mCol(i).ID = zlDatabase.GetNextId("�����ļ��ṹ")
        Next
    Case cprET_ȫ��ʾ���༭
        For i = 1 To mCol.Count
            If mCol(i).ID = 0 Then mCol(i).ID = zlDatabase.GetNextId("������������")
        Next
    Case cprET_�������༭, cprET_���������
        For i = 1 To mCol.Count
            If mCol(i).ID = 0 Then mCol(i).ID = GetEPRContentNextId
        Next
    End Select
End Sub

'################################################################################################################
'## ���ܣ�  ���������ˮ��
'################################################################################################################
Public Sub UpdateMaxKey()
    Dim i As Long, j As Long
    j = 0
    For i = 1 To mCol.Count
        If mCol(i).Key > j Then j = mCol(i).Key
    Next
    mlngKey = j
End Sub

'################################################################################################################
'## ���ܣ�  ֱ������һ���Ѿ����ڵĶ���
'##
'## ������  blnWithKeys :��ʾ�Ƿ�ͬʱ����Keyֵ
'################################################################################################################
Public Function AddExistNode(vData As cEPRElement, Optional blnWithKeys As Boolean = False) As Long
    Dim strKey As String
    If blnWithKeys = False Then
        mlngKey = mlngKey + 1
        strKey = "K" & mlngKey
    Else
        strKey = "K" & vData.Key
    End If
    mCol.Add vData.Clone(blnWithKeys), strKey
    
    If blnWithKeys = False Then
        mCol(strKey).Key = mlngKey
        AddExistNode = mlngKey
    Else
        AddExistNode = vData.Key
        Call UpdateMaxKey
    End If
End Function

'################################################################################################################
'## ���ܣ�  �����¶���
'##
'## ������  lngKey :���ڻָ��ɵļ��϶��󣬲��������Ҫ���������ˮ�š�
'################################################################################################################
Public Function Add(Optional lngKey As Long = 0) As Long
    On Error GoTo LL
    Dim objNewMember As cEPRElement
    Set objNewMember = New cEPRElement
    Dim sKey As String
    
    If lngKey = 0 Then
        mlngKey = mlngKey + 1
        sKey = "K" & mlngKey
        objNewMember.Key = mlngKey
        mCol.Add objNewMember, sKey
        Add = mlngKey
    Else
        objNewMember.Key = lngKey
        mCol.Add objNewMember, "K" & lngKey
        Call UpdateMaxKey
        Add = lngKey
    End If
    Set objNewMember = Nothing
    Exit Function
LL:
    Add = 0
End Function

'################################################################################################################
'## ���ܣ�  ��������Ҫ�ؼ���
'##
'## ������  blnWithKeys :�Ƿ���ͬKeyֵһ����
'################################################################################################################
Public Function Clone(Optional blnWithKeys As Boolean = False) As cEPRElements
    Dim cNodes As New cEPRElements, lngKey As Long
    Dim i As Long
    For i = 1 To mCol.Count
        lngKey = cNodes.Add(IIf(blnWithKeys, mCol(i).Key, 0))
        If blnWithKeys Then cNodes("K" & lngKey).Key = mCol(i).Key
        cNodes("K" & lngKey).ID = mCol(i).ID
        cNodes("K" & lngKey).�ļ�ID = mCol(i).�ļ�ID
        cNodes("K" & lngKey).��ʼ�� = mCol(i).��ʼ��
        cNodes("K" & lngKey).��ֹ�� = mCol(i).��ֹ��
        cNodes("K" & lngKey).��ID = mCol(i).��ID
        cNodes("K" & lngKey).������� = mCol(i).�������
        cNodes("K" & lngKey).�������� = mCol(i).��������
        cNodes("K" & lngKey).�����ı� = mCol(i).�����ı�
        cNodes("K" & lngKey).�Ƿ��� = mCol(i).�Ƿ���
        cNodes("K" & lngKey).����Ҫ��ID = mCol(i).����Ҫ��ID
        cNodes("K" & lngKey).�滻�� = mCol(i).�滻��
        cNodes("K" & lngKey).Ҫ������ = mCol(i).Ҫ������
        cNodes("K" & lngKey).Ҫ������ = mCol(i).Ҫ������
        cNodes("K" & lngKey).Ҫ�س��� = mCol(i).Ҫ�س���
        cNodes("K" & lngKey).Ҫ��С�� = mCol(i).Ҫ��С��
        cNodes("K" & lngKey).Ҫ�ص�λ = mCol(i).Ҫ�ص�λ
        cNodes("K" & lngKey).Ҫ�ر�ʾ = mCol(i).Ҫ�ر�ʾ
        cNodes("K" & lngKey).������̬ = mCol(i).������̬
        cNodes("K" & lngKey).Ҫ��ֵ�� = mCol(i).Ҫ��ֵ��
        cNodes("K" & lngKey).�������� = mCol(i).��������
        cNodes("K" & lngKey).���� = mCol(i).����
    Next i
    If blnWithKeys Then Call UpdateMaxKey
    Set Clone = cNodes
End Function

