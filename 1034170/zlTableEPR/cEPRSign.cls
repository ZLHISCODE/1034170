VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarKey As Long                     '�ؼ��֣��洢�ڡ������ǡ���
Private mvar���� As String                  'ǩ���û����������洢�ڡ������ı�����
Private mvarǩ����Ϣ As String              '����ǩ���ַ������洢�ڡ�Ҫ��ֵ����
Private mvarǰ������ As String              '�ƺ����߶ٺŵ�ǰ�����֣��洢�ڡ�Ҫ�����ơ���
Private mvarǩ������ As EPRSignLevel    '0,1,2,3���洢�ڡ�Ҫ�ر�ʾ����
Private mvarʱ��� As String                'ʱ������洢�ڡ�Ҫ�ص�λ����
Private mvarʱ�����Ϣ As String            '����ǩ�����ص�ʱ�������֤������Ϣ(Base64)

'����5��������ɡ��������ԡ����洢�ڡ��������ԡ��ֶ���
Private mvarǩ����ʽ As Long                '1������ǩ����2������ǩ��
Private mvarǩ������ As Long                '��1��ʼ��ţ��ɳ����Զ����亭�壨��ʾ��ͬ�汾��ǩ��������
                                            'Ŀǰ����ʹ�õ�ǩ������1
Private mvar֤��ID As Long                  '����֤��ID��
Private mvar��ʾ��ǩ As Boolean             '�Ƿ���ʾ��ǩ��־
Private mvarǩ��ʱ�� As String              'ǩ��ʱ�䣬���ڻ���ʱ�ָ����ʱ��
Private mvar��ʾʱ�� As String              '�Ƿ���ʾǩ��ʱ�估��ʾ�ĸ�ʽ��,�մ���ʾ����ʾ

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let ����(ByVal vData As String)
    mvar���� = vData
End Property

Public Property Get ����() As String
Attribute ����.VB_UserMemId = 0
    ���� = mvar����
End Property

Public Property Let ǩ����Ϣ(ByVal vData As String)
    mvarǩ����Ϣ = vData
End Property

Public Property Get ǩ����Ϣ() As String
    ǩ����Ϣ = mvarǩ����Ϣ
End Property
Public Property Let ǰ������(ByVal vData As String)
    mvarǰ������ = vData
End Property

Public Property Get ǰ������() As String
    ǰ������ = mvarǰ������
End Property

Public Property Let ǩ������(ByVal vData As EPRSignLevel)
    mvarǩ������ = vData
End Property

Public Property Get ǩ������() As EPRSignLevel
    ǩ������ = mvarǩ������
End Property

Public Property Let ǩ����ʽ(ByVal vData As Long)
    mvarǩ����ʽ = vData
End Property

Public Property Get ǩ����ʽ() As Long
    ǩ����ʽ = mvarǩ����ʽ
End Property

Public Property Let ǩ������(ByVal vData As Long)
    mvarǩ������ = vData
End Property

Public Property Get ǩ������() As Long
    ǩ������ = mvarǩ������
End Property

Public Property Let ֤��ID(ByVal vData As Long)
    mvar֤��ID = vData
End Property

Public Property Get ֤��ID() As Long
    ֤��ID = mvar֤��ID
End Property

Public Property Let ǩ��ʱ��(ByVal vData As String)
    mvarǩ��ʱ�� = vData
End Property

Public Property Get ǩ��ʱ��() As String
    ǩ��ʱ�� = mvarǩ��ʱ��
End Property

Public Property Let ��ʾʱ��(ByVal vData As String)
    mvar��ʾʱ�� = vData
End Property

Public Property Get ��ʾʱ��() As String
    ��ʾʱ�� = mvar��ʾʱ��
End Property

Public Property Let ʱ���(ByVal vData As String)
    mvarʱ��� = vData
End Property

Public Property Get ʱ���() As String
    ʱ��� = mvarʱ���
End Property

Public Property Let ʱ�����Ϣ(ByVal vData As String)
    mvarʱ�����Ϣ = vData
End Property

Public Property Get ʱ�����Ϣ() As String
    ʱ�����Ϣ = mvarʱ�����Ϣ
End Property
Public Property Let ��ʾ��ǩ(ByVal vData As Boolean)
    mvar��ʾ��ǩ = vData
End Property

Public Property Get ��ʾ��ǩ() As Boolean
    ��ʾ��ǩ = mvar��ʾ��ǩ
End Property

Public Property Let ��������(ByVal vData As String)
    Dim T As Variant
    T = Split(vData, ";")
    On Error Resume Next
    mvarǩ����ʽ = Val(T(0))
    mvarǩ������ = Val(T(1))
    mvar֤��ID = Val(T(2))
    mvar��ʾ��ǩ = (Val(T(3)) = 1)
    mvarǩ��ʱ�� = Format(T(4), "yyyy-mm-dd hh:mm:ss")
    mvar��ʾʱ�� = T(5)
    
    '��6λԤ��,ȫ�Ĳ��������ڴ��ǩ��Ҫ��Keyֵ
    
    If UBound(T) >= 7 Then
        mvarʱ�����Ϣ = CStr(T(7))
    Else
        mvarʱ�����Ϣ = ""
    End If
    Err.Clear
End Property

Public Property Get ��������() As String
'ǩ����ʽ|ǩ������|֤��ID|��ʾ��ǩ|ǩ��ʱ��|��ʾʱ��
    �������� = CStr(mvarǩ����ʽ) & ";" & CStr(mvarǩ������) & ";" & _
        CStr(mvar֤��ID) & ";" & IIf(mvar��ʾ��ǩ, 1, 0) & ";" & _
        Format(mvarǩ��ʱ��, "yyyy-mm-dd hh:mm:ss") & ";" & mvar��ʾʱ�� & ";;" & mvarʱ�����Ϣ
End Property

'################################################################################################################
'## ���ܣ�  ����ǩ������
'##
'## ������  blnWithKeys :�Ƿ���ͬKeyֵһ����
'################################################################################################################
Public Sub Clone(cNode As cTabSign)
'���ܣ�׷��ʱ�����ӳ�Ա����
    With cNode
         .���� = ""
         .ǩ����Ϣ = ""
         .ǰ������ = ""
         .ǩ������ = 0
         .�������� = Me.ǩ����ʽ & ";1;;0;;;;"
         .ʱ��� = ""
    End With
End Sub
Public Function GetSaveSQL(ByVal editType As TabET, ByVal CellTmp As cTabCell, arrSQL As Variant) As Boolean
'����:����ǩ������SQL
'˵��:�������� ��,��|ǩ����ʽ|ǩ������|֤��ID|��ʾ��ǩ|ǩ��ʱ��|��ʾʱ��
'    ǩ���û����������洢�ڡ������ı�����
'    �ƺ����߶ٺŵ�ǰ�����֣��洢�ڡ�Ҫ�����ơ���
'    ʱ������洢�ڡ�Ҫ�ص�λ����
'    ǩ�����𣬴洢�ڡ�Ҫ�ر�ʾ����
'    ����ǩ���ַ������洢�ڡ�Ҫ��ֵ����
    On Error GoTo errHand
    With CellTmp
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        Select Case editType
        Case TabET_�����ļ�����
            arrSQL(UBound(arrSQL)) = "Zl_�����ļ��ṹ_Update(" & .ID & "," & .�ļ�ID & _
                                        ",Null," & .������� & "," & .�������� & ",Null," & IIf(.��������, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|" & Me.�������� & "',Null,'" & mvar���� & "',0,Null,Null,Null" & _
                                        ",Null, Null,'" & mvarǰ������ & "',1,50,Null,'" & mvarʱ��� & "'," & mvarǩ������ & ",0,'" & mvarǩ����Ϣ & "')"
        Case TabET_ȫ��ʾ���༭
            arrSQL(UBound(arrSQL)) = "Zl_������������_Update(" & .ID & "," & .�ļ�ID & _
                                        ",Null," & .������� & "," & .�������� & ",Null," & IIf(.��������, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|" & Me.�������� & "',Null,'" & mvar���� & "',0,Null,Null,Null,Null" & _
                                        ",Null, Null,'" & mvarǰ������ & "',1,50,Null,'" & mvarʱ��� & "'," & mvarǩ������ & ",0,'" & mvarǩ����Ϣ & "')"
                                        
        Case TabET_�������༭, TabET_���������
            arrSQL(UBound(arrSQL)) = "Zl_���Ӳ�������_Update(" & .ID & "," & .�ļ�ID & "," & .��ʼ�� & "," & .��ֹ�� & _
                                        ",Null," & .������� & "," & .�������� & ",Null," & IIf(.��������, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|" & Me.�������� & "',Null,'" & mvar���� & "',0,Null,Null,Null,Null" & _
                                        ",Null, Null,'" & mvarǰ������ & "',1,50,Null,'" & mvarʱ��� & "'," & mvarǩ������ & ",0,'" & mvarǩ����Ϣ & "')"
        End Select
    End With
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function