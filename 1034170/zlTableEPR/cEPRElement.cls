VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarKey As Long                 '������
Private mvarID As Long                  '��ϱ༭������Ҫ�ص�ID,��Ҫ�ص�ID
Private mvar��ID As Long                '��ϱ༭������Ҫ�ظ�ID
Private mvar���� As String              'Ҫ�����ڵ�Ԫ�� ��|��
Private mvar�����ı� As String          '�����ı�
Private mvar�����д� As Long            '�����д�
Private mvar����Ҫ��ID As Long          '����Ҫ��ID
Private mvar�滻�� As Integer           '�Ƿ����滻��
Private mvarҪ������ As String          'Ҫ������
Private mvarҪ������ As Integer         'Ҫ������       0:��ֵ 1:���� 2:���� 3-�߼�
Private mvarҪ�س��� As Long            'Ҫ�س���
Private mvarҪ��С�� As Long            'Ҫ��С��λ��
Private mvarҪ�ص�λ As String          'Ҫ�ص�λ
Private mvarҪ�ر�ʾ As Integer         'Ҫ�ر�ʾ       0-�ı� 1-���� 2-��ѡ 3-��ѡ
Private mvar������̬ As Integer         '������̬       0-����;1-չ��(��ʾ��ѡ,��ѡ������Ҫ�ص�������̬��չ������ѡ����Ŀ)
Private mvarҪ��ֵ�� As String          'Ҫ��ֵ��       ��ֵ��Ŀ�����������Сֵ�����ֵ������������ѡ��Ŀ��ѡ����Ŀ�б�
Private mvar�������� As Boolean         '
Private mvar�Զ�ת�ı� As Boolean       '�����滻Ҫ�ض��ԣ��Ƿ�������滻���Զ�ת��Ϊ�ı�
Private mvar����   As Integer            '�Ƿ����Ҫ�ء�1-���� 0-�Ǳ���
Private mvar��̬�� As Byte              '�Ƿ������Զ��� 1-���� 0-������
Private mvar��ʼ�� As Long
Private mvar��ֹ�� As Long
Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property
Public Property Let ID(ByVal vData As Long)
    mvarID = vData
End Property

Public Property Get ID() As Long
    ID = mvarID
End Property
Public Property Let ��ID(ByVal vData As Long)
    mvar��ID = vData
End Property

Public Property Get ��ID() As Long
    ��ID = mvar��ID
End Property
Public Property Let ����(ByVal vData As String)
    mvar���� = vData
End Property
Public Property Get ����() As String
    ���� = mvar����
End Property

Public Property Let �����ı�(ByVal vData As String)
    mvar�����ı� = vData
End Property

Public Property Get �����ı�() As String
    �����ı� = mvar�����ı�
End Property
Public Property Let �����д�(ByVal vData As Long)
    mvar�����д� = vData
End Property

Public Property Get �����д�() As Long
    �����д� = mvar�����д�
End Property
Public Property Let ����Ҫ��ID(ByVal vData As Long)
    mvar����Ҫ��ID = vData
End Property

Public Property Get ����Ҫ��ID() As Long
    ����Ҫ��ID = mvar����Ҫ��ID
End Property

Public Property Let �滻��(ByVal vData As Integer)
    mvar�滻�� = vData
End Property

Public Property Get �滻��() As Integer
    �滻�� = mvar�滻��
End Property

Public Property Let Ҫ������(ByVal vData As String)
    mvarҪ������ = vData
End Property

Public Property Get Ҫ������() As String
    Ҫ������ = mvarҪ������
End Property

Public Property Let Ҫ������(ByVal vData As Integer)
    mvarҪ������ = vData
End Property

Public Property Get Ҫ������() As Integer
    Ҫ������ = mvarҪ������
End Property

Public Property Let Ҫ�س���(ByVal vData As Long)
    mvarҪ�س��� = vData
End Property

Public Property Get Ҫ�س���() As Long
    Ҫ�س��� = mvarҪ�س���
End Property

Public Property Let Ҫ��С��(ByVal vData As Long)
    mvarҪ��С�� = vData
End Property

Public Property Get Ҫ��С��() As Long
    Ҫ��С�� = mvarҪ��С��
End Property

Public Property Let Ҫ�ص�λ(ByVal vData As String)
    mvarҪ�ص�λ = vData
End Property

Public Property Get Ҫ�ص�λ() As String
    Ҫ�ص�λ = mvarҪ�ص�λ
End Property

Public Property Let Ҫ�ر�ʾ(ByVal vData As Integer)
    mvarҪ�ر�ʾ = vData
End Property

Public Property Get Ҫ�ر�ʾ() As Integer
    Ҫ�ر�ʾ = mvarҪ�ر�ʾ
End Property

Public Property Let ������̬(ByVal vData As Integer)
    mvar������̬ = vData
End Property

Public Property Get ������̬() As Integer
    ������̬ = mvar������̬
End Property

Public Property Let Ҫ��ֵ��(ByVal vData As String)
    mvarҪ��ֵ�� = vData
End Property

Public Property Get Ҫ��ֵ��() As String
    Ҫ��ֵ�� = mvarҪ��ֵ��
End Property
Public Property Let ��������(ByVal vData As Boolean)
    mvar�������� = vData
End Property

Public Property Get ��������() As Boolean
    �������� = mvar��������
End Property
Public Property Let �Զ�ת�ı�(ByVal vData As Boolean)
    mvar�Զ�ת�ı� = vData
End Property

Public Property Get �Զ�ת�ı�() As Boolean
    �Զ�ת�ı� = mvar�Զ�ת�ı�
End Property
Public Property Let ����(ByVal vData As Integer)
    mvar���� = vData
End Property
Public Property Get ����() As Integer
    ���� = mvar����
End Property
Public Property Let ��̬��(ByVal vData As Byte)
    mvar��̬�� = vData
End Property
Public Property Get ��̬��() As Byte
    ��̬�� = mvar��̬��
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
    ��ֹ�� = IIf(mvar��ֹ�� = 0, 0, mvar��ֹ��)
End Property
Public Property Get ��������() As String
    �������� = IIf(mvar�Զ�ת�ı�, 1, 0) & "|" & mvar���� & "|" & mvar��̬��
End Property
Public Property Let ��������(ByRef vData As String)
Dim strFirst As String, strData As String
    strFirst = Split(vData, "|")(0)
    If InStr(strFirst, ",") > 0 Then
        strData = Mid(vData, InStr(vData, "|") + 1)
    Else
        strData = vData
    End If
        
    Select Case UBound(Split(strData, "|"))
        Case Is <= 0
            mvar�Զ�ת�ı� = IIf(Val(strData) = 0, False, True)
            mvar���� = 0
            mvar��̬�� = 0
        Case 1
            mvar�Զ�ת�ı� = IIf(Val(Split(strData, "|")(0)) = 0, False, True)
            mvar���� = Val(Split(strData, "|")(1))
            mvar��̬�� = 0
        Case 2
            mvar�Զ�ת�ı� = IIf(Val(Split(strData, "|")(0)) = 0, False, True)
            mvar���� = Val(Split(strData, "|")(1))
            mvar��̬�� = Val(Split(strData, "|")(2))
    End Select
End Property

'################################################################################################################
'## ���ܣ�  ��������Ҫ�ض���
'##
'## ������  blnWithKeys :�Ƿ���ͬKeyֵһ����
'################################################################################################################
Public Sub Clone(cNode As cTabElement)
'����
    With cNode
        .�����ı� = Me.�����ı�
        .�����д� = Me.�����д�
        .����Ҫ��ID = Me.����Ҫ��ID
        .�滻�� = Me.�滻��
        .Ҫ������ = Me.Ҫ������
        .Ҫ������ = Me.Ҫ������
        .Ҫ�س��� = Me.Ҫ�س���
        .Ҫ��С�� = Me.Ҫ��С��
        .Ҫ�ص�λ = Me.Ҫ�ص�λ
        .Ҫ�ر�ʾ = Me.Ҫ�ر�ʾ
        .������̬ = Me.������̬
        .Ҫ��ֵ�� = Me.Ҫ��ֵ��
        .�������� = Me.��������
        .�������� = Me.��������
        .���� = Me.����
        .��̬�� = Me.��̬��
        .��ʼ�� = Me.��ʼ��
        .��ֹ�� = Me.��ֹ��
    End With
End Sub
'################################################################################################################
'## ���ܣ�  �������Ҫ�ص�ԭʼ�����Ƿ���ڣ�����XML����ʱ����֤��
'################################################################################################################
Public Function CheckValid() As Boolean
    Dim rs As New Recordset
    gstrSQL = "Select Count(*) From ����������Ŀ Where ID = [1] And ������ = [2]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cTabElement", Me.����Ҫ��ID, Me.Ҫ������)
    If rs.EOF Then
        CheckValid = False
    Else
        CheckValid = (rs(0) > 0)
    End If
End Function
Public Function GetSaveSQL(ByVal editType As TabET, ByVal CellTmp As cTabCell, arrSQL As Variant) As Boolean
'����:���ɵ�Ҫ�غͻ��������Ҫ�ر���SQL,CellTmp��
    On Error GoTo errHand
    With CellTmp
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        Select Case editType
        Case TabET_�����ļ�����
            arrSQL(UBound(arrSQL)) = "Zl_�����ļ��ṹ_Update(" & IIf(mvar��ID = 0, .ID, mvarID) & "," & .�ļ�ID & "," & _
                                        IIf(.�������� = cprCTElement, "Null", .ID) & "," & .������� & "," & cprCTElement & ",Null," & _
                                        IIf(.��������, 1, 0) & ",'" & .Row & "," & .Col & "|" & Me.�������� & "'," & mvar�����д� & ",'" & _
                                        IIf(.�������� = cprCTElement, .�����ı�, mvar�����ı�) & "',0,Null,Null,Null," & _
                                        IIf(mvar����Ҫ��ID = 0, "Null", mvar����Ҫ��ID) & "," & mvar�滻�� & ",'" & mvarҪ������ & "'," & mvarҪ������ & "," & _
                                        mvarҪ�س��� & "," & mvarҪ��С�� & ",'" & mvarҪ�ص�λ & "'," & mvarҪ�ر�ʾ & "," & mvar������̬ & ",'" & mvarҪ��ֵ�� & "')"
        Case TabET_ȫ��ʾ���༭
            arrSQL(UBound(arrSQL)) = "zl_������������_update(" & IIf(mvar��ID = 0, .ID, mvarID) & "," & .�ļ�ID & "," & _
                                        IIf(.�������� = cprCTElement, "Null", .ID) & "," & .������� & "," & cprCTElement & ",Null," & _
                                        IIf(.��������, 1, 0) & ",'" & .Row & "," & .Col & "|" & Me.�������� & "'," & mvar�����д� & ",'" & _
                                        IIf(.�������� = cprCTElement, .�����ı�, mvar�����ı�) & "',0,Null,Null,Null,Null," & _
                                        IIf(mvar����Ҫ��ID = 0, "Null", mvar����Ҫ��ID) & "," & mvar�滻�� & ",'" & mvarҪ������ & "'," & mvarҪ������ & "," & _
                                        mvarҪ�س��� & "," & mvarҪ��С�� & ",'" & mvarҪ�ص�λ & "'," & mvarҪ�ر�ʾ & "," & mvar������̬ & ",'" & mvarҪ��ֵ�� & "')"
        Case TabET_�������༭, TabET_���������
            arrSQL(UBound(arrSQL)) = "Zl_���Ӳ�������_Update(" & IIf(mvar��ID = 0, .ID, mvarID) & "," & .�ļ�ID & "," & mvar��ʼ�� & "," & mvar��ֹ�� & "," & _
                                        IIf(.�������� = cprCTElement, "Null", .ID) & "," & .������� & "," & cprCTElement & ",Null," & _
                                        IIf(.��������, 1, 0) & ",'" & .Row & "," & .Col & "|" & Me.�������� & "'," & mvar�����д� & ",'" & _
                                        IIf(.�������� = cprCTElement, .�����ı�, mvar�����ı�) & "',0,Null,Null,Null,Null," & _
                                        IIf(mvar����Ҫ��ID = 0, "Null", mvar����Ҫ��ID) & "," & mvar�滻�� & ",'" & mvarҪ������ & "'," & mvarҪ������ & "," & _
                                        mvarҪ�س��� & "," & mvarҪ��С�� & ",'" & mvarҪ�ص�λ & "'," & mvarҪ�ر�ʾ & "," & mvar������̬ & ",'" & mvarҪ��ֵ�� & "')"
        End Select
    End With
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function InsertIntoEditor(ByRef edtThis As Object, ByVal editType As Integer, Optional ByVal lngStartPos As Long = -1) As Boolean
'******************************************************************************************************************
' ���ܣ�  ��������Ҫ�ص�Editor��
'
' ������  edtThis         :��ǰ�ı༭���ؼ�
'         lngStartPos     :��ǰλ��
'         blnFirstIns     :�Ƿ��ǵ�һ�β��루Ĭ�ϵ�һ�β������ĩβ�Ӷ��ţ�
'******************************************************************************************************************
Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    
    If lngStartPos = -1 Then lngStartPos = edtThis.Selection.StartPos
    bInKeys = IsBetweenAnyKeys(edtThis, lngStartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then InsertIntoEditor = False: Exit Function    '��֤���ܲ���ؼ����ڲ�
    
    If edtThis.Range(lngStartPos, lngStartPos).Font.ForeColor = PROTECT_FORECOLOR And edtThis.Range(lngStartPos, lngStartPos).Font.Protected Then Exit Function

    Dim p As Long, lngKey As Long, blnForce As Boolean, lngLen As Long, lngLen2 As Long, strTmp As String
    p = lngStartPos
    blnForce = edtThis.ForceEdit

    With edtThis
        .Freeze
        .ForceEdit = True
        lngKey = Me.Key
        
        strTmp = Format(lngKey, "00000000") & "," & IIf(Me.��������, 1, 0) & ",0)"

        .Range(lngStartPos, lngStartPos).Font.Protected = False
        .Range(lngStartPos, lngStartPos).Font.Hidden = False
        
        If Me.�滻�� = 1 And (editType = TabET_�������༭ Or editType = TabET_���������) Then '�Զ��滻Ҫ����"�Զ�ת�ı�"����
            If Trim(Me.�����ı�) = "" Then
                If Me.�Զ�ת�ı� Then
                    lngLen = Len(" ")
                    lngLen2 = Len(Me.Ҫ�ص�λ)
                    .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & " " & Me.Ҫ�ص�λ & "EE(" & strTmp
                Else
                    lngLen = Len("[" & Me.Ҫ������ & "]")
                    lngLen2 = Len(Me.Ҫ�ص�λ)
                    .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & "[" & Me.Ҫ������ & "]" & Me.Ҫ�ص�λ & "EE(" & strTmp
                End If
            Else
                lngLen = Len(Me.�����ı�)
                lngLen2 = Len(Me.Ҫ�ص�λ)
                .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & Me.�����ı� & Me.Ҫ�ص�λ & "EE(" & strTmp
            End If
        Else
            If Me.������̬ = 0 Then
                lngLen = Len(IIf(Me.�����ı� = "", "[" & Me.Ҫ������ & "]", Me.�����ı�))
                lngLen2 = Len(Me.Ҫ�ص�λ)
                .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & IIf(Me.�����ı� = "", "[" & Me.Ҫ������ & "]", Me.�����ı�) & Me.Ҫ�ص�λ & "EE(" & strTmp
            Else
                lngLen = Len(Me.�����ı�)
                lngLen2 = Len(Me.Ҫ�ص�λ)
                .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & Me.�����ı� & Me.Ҫ�ص�λ & "EE(" & strTmp
            End If
        End If
        .Range(lngStartPos, lngStartPos + 32 + lngLen + lngLen2).Font.Protected = True '����Ҫ�ؼ���������
        .Range(lngStartPos, lngStartPos + 16).Font.Hidden = True                       '����ǰ���ؼ���
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen + lngLen2).Font.Hidden = False   '������ʾ
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen + lngLen2).Font.BackColor = ELE_BACKCOLOR
        .Range(lngStartPos + 16 + lngLen + lngLen2, lngStartPos + 32 + lngLen + lngLen2).Font.Hidden = True '����β���ؼ���
        lngStartPos = lngStartPos + 32 + lngLen + lngLen2                                   'λ�ú���
        .Range(lngStartPos, lngStartPos).Selected                                           '��궨λ
        .ForceEdit = blnForce
        .UnFreeze
    End With
    InsertIntoEditor = True
End Function

'################################################################################################################
'## ���ܣ�  ˢ����ʾ
'##
'## ������  edtThis         :��ǰ�ı༭���ؼ�
'################################################################################################################
Public Function Refresh(ByRef edtThis As Object) As Boolean
    '��������Ҫ�ر༭���
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "E", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True
            Dim ParaFmt As New cTabParaFormat, FontFmt As New cTabFontFormat
            Set ParaFmt = edtThis.Range(lKSE, lKSE).Para.GetParaFmt
            Set FontFmt = edtThis.Range(lKSE, lKSE).Font.GetFontFmt
            If FontFmt.Underline = cprwave Then FontFmt.Underline = cprnone
            
            .Range(lKSS, lKEE).Text = ""
            Me.InsertIntoEditor edtThis, lKSS
            .Range(lKSE, lKSE + Len(Me.�����ı�)).Selected
            .Range(lKSE, lKES).Font.SetFontFmt FontFmt
            .Range(lKSE, lKSE).Para.SetParaFmt ParaFmt
            .ForceEdit = blnForce
            .UnFreeze
        End With
        Refresh = True
    Else
        Refresh = False
    End If
End Function

'################################################################################################################
'## ���ܣ�  �ӱ༭����ɾ��
'##
'## ������  edtThis         :��ǰ�ı༭���ؼ�
'################################################################################################################
Public Function DeleteFromEditor(ByRef edtThis As Object) As Boolean
    '��������Ҫ�ر༭���
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "E", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True
            .Range(lKSS, lKSS).Selected
            .Range(lKSS, lKEE).Text = ""
            .ForceEdit = blnForce
            .UnFreeze
        End With
        DeleteFromEditor = True
    Else
        DeleteFromEditor = False
    End If
End Function
'################################################################################################################
'## ���ܣ�  ������ж����ID�Ƿ���Ч
'##
'## ������  eEditType :��ǰ�ı༭ģʽ
'##
'## ˵����  ���û��ID�������ݿ�����ȡһ��ΨһID�š���ID=0������ǵ�Ԫ����CellȡID
'################################################################################################################
Public Sub CheckValidID(eEditType As TabET)
    Dim i As Long, strTable As String
    Select Case eEditType
    Case TabET_�����ļ�����
        strTable = "�����ļ��ṹ"
    Case TabET_ȫ��ʾ���༭
        strTable = "������������"
    Case TabET_�������༭, TabET_���������
        strTable = "���Ӳ�������"
    End Select
    
    If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId(strTable)
End Sub