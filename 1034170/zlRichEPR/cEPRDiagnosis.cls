VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarKey As Long         '������
Private mvarID As Double          'ID
Private mvar�ļ�ID As Long      '�ļ�ID
Private mvar��ID As Double        '��ID��������������
Private mvar������� As Double    '�������

Private mvar���� As Long        '������ͣ�11-�������;21-��Ժ�������;22-��Ժȷ�����;23-��Ժ�������;24-��Ժ������ϣ�31-��Ժ���;41-��ǰ���;42-�������
Private mvar��ҽ As Byte        '�Ƿ���ҽ��0-��ҽ;1-��ҽ
Private mvar����id As Long      '��Ӧ��������Ŀ¼(ICD����ҽ����)��ID
Private mvar���id As Long      '��Ӧ�������Ŀ¼��ID
Private mvar֤��id As Long      '��Ӧ�������Ŀ¼��ID
Private mvar���� As String      '������ϵ���������ҽ��ϸ�ʽΪ����(֤)�����磺�з�(̵�ȸ�ʵ����̵����)
Private mvar���� As Byte        '�Ƿ����0-ȷ��;1-����
Private mvar���� As String
Private mvar��Ժ���� As String  '��Ժ���ͬ����ҳʱʹ��
Private mvar��Ժ��� As String  '��Ժ���ͬ����ҳʱʹ��

Private mvar��ʼ�� As Long      '��ʼ�棬ר���ڡ����Ӳ�����¼��
Private mvar��ֹ�� As Long      '��ֹ�棬ר���ڡ����Ӳ�����¼��

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let ID(ByVal vData As Double)
    mvarID = vData
End Property

Public Property Get ID() As Double
    ID = mvarID
End Property

Public Property Let �ļ�ID(ByVal vData As Long)
    mvar�ļ�ID = vData
End Property

Public Property Get �ļ�ID() As Long
    �ļ�ID = mvar�ļ�ID
End Property

Public Property Let ��ID(ByVal vData As Double)
    mvar��ID = vData
End Property

Public Property Get ��ID() As Double
    ��ID = mvar��ID
End Property

Public Property Let �������(ByVal vData As Double)
    mvar������� = vData
End Property

Public Property Get �������() As Double
    ������� = mvar�������
End Property

Public Property Let ����(ByVal vData As Long)
    mvar���� = vData
End Property

Public Property Get ����() As Long
    ���� = mvar����
End Property

Public Property Let ��ҽ(ByVal vData As Byte)
    mvar��ҽ = vData
End Property

Public Property Get ��ҽ() As Byte
    ��ҽ = mvar��ҽ
End Property

Public Property Let ����id(ByVal vData As Long)
    mvar����id = vData
End Property

Public Property Get ����id() As Long
    ����id = mvar����id
End Property

Public Property Let ���id(ByVal vData As Long)
    mvar���id = vData
End Property

Public Property Get ���id() As Long
    ���id = mvar���id
End Property

Public Property Let ֤��id(ByVal vData As Long)
    mvar֤��id = vData
End Property

Public Property Get ֤��id() As Long
    ֤��id = mvar֤��id
End Property

Public Property Let ����(ByVal vData As String)
    mvar���� = vData
End Property

Public Property Get ����() As String
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As Byte)
    mvar���� = vData
End Property

Public Property Get ����() As Byte
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As String)
    mvar���� = vData
End Property

Public Property Get ����() As String
    ���� = mvar����
End Property

Public Property Let ��Ժ����(ByVal vData As String)
    mvar��Ժ���� = vData
End Property

Public Property Get ��Ժ����() As String
    ��Ժ���� = mvar��Ժ����
End Property
Public Property Let ��Ժ���(ByVal vData As String)
    mvar��Ժ��� = vData
End Property

Public Property Get ��Ժ���() As String
    ��Ժ��� = mvar��Ժ���
End Property

Public Property Let ��ʼ��(ByVal vData As Long)
    mvar��ʼ�� = vData
End Property

Public Property Get ��ʼ��() As Long
    ��ʼ�� = IIf(mvar��ʼ�� = 0, 1, mvar��ʼ��)
End Property

Public Property Let ��ֹ��(ByVal vData As Long)
    mvar��ֹ�� = vData
End Property

Public Property Get ��ֹ��() As Long
    ��ֹ�� = mvar��ֹ��
End Property


'################################################################################################################
'## ���ܣ�  ���ƶ���
'##
'## ������  blnWithKeys :�Ƿ���ͬKeyֵһ����
'################################################################################################################
Public Function Clone(Optional blnWithKey As Boolean = False) As cEPRDiagnosis
    '���ڿ������ݣ�����һ���������
    Dim cNode As New cEPRDiagnosis
    If blnWithKey Then cNode.Key = Me.Key
    cNode.ID = Me.ID
    cNode.�ļ�ID = Me.�ļ�ID
    cNode.��ʼ�� = Me.��ʼ��
    cNode.��ֹ�� = Me.��ֹ��
    cNode.��ID = Me.��ID
    cNode.������� = Me.�������
    
    cNode.���� = Me.����
    cNode.��ҽ = Me.��ҽ
    cNode.����id = Me.����id
    cNode.���id = Me.���id
    cNode.֤��id = Me.֤��id
    cNode.���� = Me.����
    cNode.���� = Me.����
    cNode.���� = Me.����
    cNode.��Ժ���� = Me.��Ժ����
    cNode.��Ժ��� = Me.��Ժ���

    Set Clone = cNode
End Function
Public Function FillDiagnosisMember(ByVal rs As ADODB.Recordset, ByVal strTableName As String) As Boolean
'���ܣ������ݼ�¼������Ա
Dim aryProperty() As String

    If rs Is Nothing Then Exit Function
    If rs.EOF Then Exit Function
    
    If NVL(rs!������, 0) <> 0 Then Me.Key = Val("" & rs!������)
    Me.ID = rs!ID
    Me.�ļ�ID = Val("" & rs!�ļ�ID)
    Me.��ID = Val("" & rs!��ID)
    Me.������� = Val("" & rs!�������)
    Me.���� = "" & rs!�����ı�
    aryProperty = Split("" & rs!��������, ";")
    Me.���� = Val(aryProperty(0))
    Me.��ҽ = Val(aryProperty(1))
    Me.����id = Val(aryProperty(2))
    Me.���id = Val(aryProperty(3))
    Me.֤��id = Val(aryProperty(4))
    Me.���� = Val(aryProperty(5))
    Me.���� = Format(aryProperty(6), "yyyy-mm-dd hh:mm:ss")
    If UBound(aryProperty) > 6 Then
        Me.��Ժ���� = aryProperty(7)
        Me.��Ժ��� = aryProperty(8)
    End If

    If InStr(strTableName, "���Ӳ�������") > 0 Then '�����ļ��ṹ�Ͳ�����������û���������ֶ�,���������������
        Me.��ʼ�� = NVL(rs!��ʼ��, 1)
        Me.��ֹ�� = NVL(rs!��ֹ��, 0)
    End If

    FillDiagnosisMember = True
End Function

'################################################################################################################
'## ���ܣ�  �������ID�Ƿ���Ч
'##
'## ������  eEditType :��ǰ�ı༭ģʽ
'##
'## ˵����  ���û��ID�������ݿ�����ȡһ��ΨһID�š�
'################################################################################################################
Public Sub CheckValidID(eEdtType As EditTypeEnum)
    Select Case eEdtType
    Case cprET_�����ļ�����
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("�����ļ��ṹ")
    Case cprET_ȫ��ʾ���༭
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("������������")
    Case cprET_�������༭, cprET_���������
        If Me.ID = 0 Then Me.ID = GetEPRContentNextId
    End Select
End Sub

'################################################################################################################
'## ���ܣ�  ����ָ������Editor��
'##
'## ������  edtThis         :��ǰ�ı༭���ؼ�
'##         lngStartPos     :��ǰλ��
'##         blnForceInsert  :�Ƿ�ǿ�Ʋ��룬���޸�Ҫ��ʱ���ǰ�󶼱�������ʱ��Ҫǿ�Ʋ���
'################################################################################################################
Public Function InsertIntoEditor(ByRef edtThis As Object, _
    Optional ByVal lngStartPos As Long = -1, _
    Optional ByVal blnForceInsert As Boolean = False) As Boolean
    
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim blnForce As Boolean, lngLen As Long, strTmp As String, strText As String
    
    If lngStartPos = -1 Then lngStartPos = edtThis.Selection.StartPos
    bInKeys = IsBetweenAnyKeys(edtThis, lngStartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then InsertIntoEditor = False: Exit Function    '��֤���ܲ���ؼ����ڲ�
    If blnForceInsert = False And edtThis.Range(lngStartPos, lngStartPos).Font.ForeColor = PROTECT_FORECOLOR _
        And edtThis.Range(lngStartPos, lngStartPos).Font.Protected Then Exit Function
    If Trim(Me.����) = "" Then Exit Function  '���ǩ������Ϊ�գ�Ҳ�˳�
    
    With edtThis
        .Freeze
        blnForce = .ForceEdit
        .ForceEdit = True
        
        strTmp = Format(Me.Key, "00000000") & ",0,0)"

        .Range(lngStartPos, lngStartPos).Font.Protected = False
        .Range(lngStartPos, lngStartPos).Font.Hidden = False
        
        strText = Me.����
        lngLen = Len(strText)
        .Range(lngStartPos, lngStartPos) = "DS(" & strTmp & strText & "DE(" & strTmp
            
        .Range(lngStartPos, lngStartPos + 32 + lngLen).Font.Protected = True
        .Range(lngStartPos, lngStartPos + 16).Font.Hidden = True
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Hidden = False
        'ɾ����
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Strikethrough = (Me.��ֹ�� > 0)
        '�»���
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Underline = cprNone
        'ǰ��ɫ
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.ForeColor = GetCharColor(Me.��ʼ��, Me.��ֹ�� + 1)
        '����ɫ
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.BackColor = ELE_BACKCOLOR
        .Range(lngStartPos + 16 + lngLen, lngStartPos + 32 + lngLen).Font.Hidden = True
        lngStartPos = lngStartPos + 32 + lngLen
        .Range(lngStartPos, lngStartPos).Selected

        .ForceEdit = blnForce
        .UnFreeze
    End With
    InsertIntoEditor = True
End Function

'################################################################################################################
'## ���ܣ�  ˢ�¶������ʾ
'##
'## ������  edtThis         :��ǰ�ı༭���ؼ�
'################################################################################################################
Public Function Refresh(ByRef edtThis As Object) As Boolean
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "D", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True
            Dim ParaFmt As New cParaFormat
            Set ParaFmt = edtThis.Range(lKSE, lKES).Para.GetParaFmt

            .Range(lKSS, lKEE) = ""
            Me.InsertIntoEditor edtThis, lKSS, True
            .Range(lKSE, lKSE + Len(Me.����)).Selected
            .Range(lKSE, lKES).Para.SetParaFmt ParaFmt
            .ForceEdit = blnForce
            .UnFreeze
        End With
        Refresh = True
    Else
        Refresh = False
    End If
End Function

'################################################################################################################
'## ���ܣ�  �ӱ༭����ɾ������
'##
'## ������  edtThis         :��ǰ�ı༭���ؼ�
'################################################################################################################
Public Function DeleteFromEditor(ByRef edtThis As Object) As Boolean
    '��������Ҫ�ر༭���
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "D", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True
            .Range(lKSS, lKEE) = ""
            .ForceEdit = blnForce
            .UnFreeze
        End With
        DeleteFromEditor = True
    Else
        DeleteFromEditor = False
    End If
End Function

'################################################################################################################
'## ���ܣ�  ��ȡ���浽���ݿ��SQL���
'##
'## ������  eEdtType    :��ǰ�༭��ʽ��
'## ����ֵ������ �洢���̵��õ�SQL��䡣
'################################################################################################################
Public Function GetSaveSQL(ByVal eEdtType As EditTypeEnum, ByRef ArraySQL() As String) As Boolean
    
    Dim lngLBound As Long, lngUBound As Long    '�����������С����±�
    Err = 0: On Error Resume Next
    lngLBound = LBound(ArraySQL): lngUBound = UBound(ArraySQL)
    If Err <> 0 Then lngLBound = 0: lngUBound = -1
    ReDim Preserve ArraySQL(lngLBound To lngUBound + 1)

    Dim strProperty As String
    strProperty = Me.���� & ";" & Me.��ҽ & ";" & Me.����id & ";" & Me.���id & ";" & Me.֤��id & ";" & Me.���� & ";" & Format(Me.����, "yyyy-mm-dd hh:mm:ss") & ";" & Me.��Ժ���� & ";" & Me.��Ժ���
    
    Call CheckValidID(eEdtType)         '���ID��Ч�ԣ�
    Select Case eEdtType
    Case cprET_�����ļ�����
        ArraySQL(lngUBound + 1) = "Zl_�����ļ��ṹ_Update(" & Me.ID & "," & Me.�ļ�ID & "," & _
            IIf(Me.��ID = 0, "NULL", Me.��ID) & "," & Me.������� & ",7," & Me.Key & ",0,'" & strProperty & "',NULL,'" & Me.���� & "')"
    Case cprET_ȫ��ʾ���༭
        ArraySQL(lngUBound + 1) = "Zl_������������_Update(" & Me.ID & "," & Me.�ļ�ID & "," & _
            IIf(Me.��ID = 0, "NULL", Me.��ID) & "," & Me.������� & ",7," & Me.Key & ",0,'" & strProperty & "',NULL,'" & Me.���� & "')"
    Case cprET_�������༭, cprET_���������
        ArraySQL(lngUBound + 1) = "ZL_���Ӳ�������_Update(" & Me.ID & "," & Me.�ļ�ID & "," & Me.��ʼ�� & "," & Me.��ֹ�� & "," & _
            IIf(Me.��ID = 0, "NULL", Me.��ID) & "," & Me.������� & ",7," & Me.Key & ",0,'" & strProperty & "',NULL,'" & Me.���� & "')"
    End Select
    GetSaveSQL = True
End Function

'################################################################################################################
'## ���ܣ�  ѡ�иö���
'##
'## ������  edtThis         :��ǰ�ı༭���ؼ�
'################################################################################################################
Public Function Selected(ByRef edtThis As Object) As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    On Error Resume Next
    bFinded = FindKey(edtThis, "D", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Range(lKSE, lKES).Selected
        End With
        Selected = True
    Else
        Selected = False
    End If
End Function
