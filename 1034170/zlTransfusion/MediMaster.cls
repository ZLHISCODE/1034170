VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MediBill"
Attribute VB_Ext_KEY = "Member0" ,"MediBill"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'��������ֵ�ľֲ�����
Option Explicit

Private mvarNO As String '�ֲ�����
Private mvar�Һŵ� As String '�ֲ�����
Private mvar����� As String
Private mvar���� As String '�ֲ�����
Private mvar�Ա� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar����Ա As String '�ֲ�����
Private mvar�Ǽ�ʱ�� As Date '�ֲ�����
Private mvar����ʱ�� As Date '�ֲ�����
Private mvar�ϼ� As Currency '�ֲ�����
Private mvarժҪ As String '�ֲ�����
Private mvar����ID As Long
Private mvar����ID As Long
Private mvar�������� As String

Private mColBills As Collection  '����Bill

Public Property Let ժҪ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ժҪ = 5
    mvarժҪ = MidUni(vData, 1, 200)
End Property

Public Property Get ժҪ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ժҪ
    ժҪ = mvarժҪ
End Property

Public Property Let �ϼ�(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ϼ� = 5
    mvar�ϼ� = vData
End Property


Public Property Get �ϼ�() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�ϼ�
    �ϼ� = mvar�ϼ�
End Property


Public Property Let �Ǽ�ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ǽ�ʱ�� = 5
    mvar�Ǽ�ʱ�� = vData
End Property


Public Property Get �Ǽ�ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�Ǽ�ʱ��
    �Ǽ�ʱ�� = mvar�Ǽ�ʱ��
End Property

Public Property Let ����ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ǽ�ʱ�� = 5
    mvar����ʱ�� = vData
End Property


Public Property Get ����ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�Ǽ�ʱ��
    ����ʱ�� = mvar����ʱ��
End Property


Public Property Let ����Ա(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����Ա = 5
    mvar����Ա = vData
End Property


Public Property Get ����Ա() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����Ա
    ����Ա = mvar����Ա
End Property



Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    ���� = mvar����
End Property


Public Property Let �Ա�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ա� = 5
    mvar�Ա� = vData
End Property


Public Property Get �Ա�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�Ա�
    �Ա� = mvar�Ա�
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    ���� = mvar����
End Property

Public Property Let ��������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar�������� = vData
End Property

Public Property Get ��������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    �������� = mvar��������
End Property

Public Property Let �Һŵ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�

    mvar�Һŵ� = vData
End Property

Public Property Get �Һŵ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�

    �Һŵ� = mvar�Һŵ�
End Property

Public Property Let �����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����� = 5
    mvar����� = vData
End Property

Public Property Get �����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�����
    ����� = mvar�����
End Property

Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�

    mvar����ID = vData
End Property

Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�

    ����ID = mvar����ID
End Property

Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�

    mvar����ID = vData
End Property

Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�

    ����ID = mvar����ID
End Property

Public Property Let NO(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.NO = 5
    mvarNO = vData
End Property

Public Property Get NO() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.NO
    NO = mvarNO
End Property

Public Property Get BillItem(vntIndexKey As Variant) As MediBill
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set BillItem = mColBills(vntIndexKey)
End Property

Public Property Get BillCount() As Long
    ' x.Count
    BillCount = mColBills.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)
    mColBills.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    '������������ For...Each �﷨ö�ٸü��ϡ�
    Set NewEnum = mColBills.[_NewEnum]
End Property

Private Sub Class_Initialize()
    '������󴴽�����
    Set mColBills = New Collection
End Sub

Private Sub Class_Terminate()
    '����ֹ���ƻ�����
    Set mColBills = Nothing
End Sub

Public Sub AddBill(ByVal objMediBill As MediBill, Optional ByVal StrKey As String)
    '�����¶���
    Dim objNewMember As MediBill
    
    Set objNewMember = objMediBill
    '���ô��뷽��������
    If Len(StrKey) = 0 Then
        mColBills.Add objNewMember
    Else
        mColBills.Add objNewMember, StrKey
    End If
    Set objNewMember = Nothing
    
End Sub

Public Function GetNextNo() As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    GetNextNo = zlDatabase.GetNextNo(19)
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub InsertBill(ByVal strNO As String, ByVal dateCurr As Date)
    Dim strInsSQL As String, strSQL As String, rsTmp As New ADODB.Recordset
    Dim lng���ID As Long, lng��ִ������ As Long, lng�������� As Long, date����ʱ�� As Date
    Dim objBIll As MediBill
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    For Each objBIll In mColBills
        With objBIll
        
            '---- ҽ������ʱ,Ҫ����ִ�е�����ͬ��д��,�ݴ��
'            If .ҽ��ID <> 0 And .ҩƷID <> 0 And .���ϵ�� = 1 Then
'                strSQL = "Select A.���id,B.����ʱ��, (B.�������� / C.����ϵ��) As ����" & vbNewLine & _
'                        "From ҩƷ��� C, ����ҽ������ B, ����ҽ����¼ A" & vbNewLine & _
'                        "Where A.�շ�ϸĿid = C.ҩƷid And A.ID = B.ҽ��id And B.���ͺ� = [2] And B.ҽ��id = [1]"
'                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "InsertBill", .ҽ��ID, .���ͺ�)
'                If Not rsTmp.EOF Then
'                    lng���ID = Val(rsTmp.Fields("���ID"))
'                    lng�������� = Val(rsTmp.Fields("����"))
'                    date����ʱ�� = CDate(rsTmp.Fields("����ʱ��"))
'                End If
'
'                If lng���ID <> 0 And lng�������� > 0 Then
'                    strSQL = "Select B.��������, Sum(Nvl(A.��������, 0)) As ��ִ������" & vbNewLine & _
'                            "From ����ҽ������ B, ����ҽ��ִ�� A" & vbNewLine & _
'                            "Where A.ҽ��id(+) = B.ҽ��id And A.���ͺ�(+) = B.���ͺ� And B.ҽ��id = [1] And B.���ͺ� = [2]" & vbNewLine & _
'                            "Group By B.��������"
'                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "InsertBill", lng���ID, .���ͺ�)
'
'                    If Not rsTmp.EOF Then
'                        '�������������
'                        lng��ִ������ = (Val(rsTmp.Fields("��ִ������")) / Val(rsTmp.Fields("��������"))) * lng��������
'                        If lng��ִ������ > .���� Then
'                            lng��ִ������ = .����
'                        End If
'                        If lng��ִ������ > 0 Then
'                            .ʹ��״̬ = 1
'                            strInsSQL = "ZL_�ݴ�ҩƷ��¼_Insert('" & strNO & "'," & .��� & "," & _
'                                     mvar����ID & "," & .ҽ��ID & "," & .���ͺ� & "," & .ҩƷID & ",'" & _
'                                     .ҩƷ���� & "','" & .��� & "'," & .ִ�з��� & "," & _
'                                     .ʹ��״̬ & ",'" & "�Զ�д��" & "'," & -1 & ",'" & _
'                                     .���㵥λ & "'," & .���� & "," & lng��ִ������ & "," & _
'                                     .���� & "," & lng��ִ������ * .���� & ",'" & mvar����Ա & "'," & _
'                                     mvar����ID & ",to_date('" & Format(date����ʱ��, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'))"
'                            Call zlDatabase.ExecuteProcedure(strInsSQL, "MediMaster.InsertBill")
'                        End If
'                    End If
'                End If
'            End If
            '-- д��Ĵ��¼
            strInsSQL = "ZL_�ݴ�ҩƷ��¼_Insert('" & strNO & "'," & .��� & "," & _
                     mvar����ID & "," & IIf(.ҽ��ID = 0, "Null", .ҽ��ID) & "," & IIf(.���ͺ� = 0, "Null", .���ͺ�) & "," & .ҩƷID & ",'" & _
                     .ҩƷ���� & "','" & .��� & "'," & .ִ�з��� & "," & _
                     .ʹ��״̬ & ",'" & mvarժҪ & "'," & .���ϵ�� & ",'" & _
                     .���㵥λ & "'," & .���� & "," & .���� & "," & _
                     .���� & "," & .��� & ",'" & mvar����Ա & "'," & _
                     mvar����ID & ",to_date('" & Format(dateCurr, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'))"
            Call zlDatabase.ExecuteProcedure(strInsSQL, "MediMaster.InsertBill")
        End With
    Next
    gcnOracle.CommitTrans
    Exit Sub
errHandle:
    
    Call ErrCenter
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Public Property Get ʹ�����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�

    Dim objBIll As MediBill
    Dim intʹ�� As Integer
    For Each objBIll In mColBills
        If objBIll.���ϵ�� = -1 Then
            intʹ�� = intʹ�� + 1
        End If
    Next
    
    If intʹ�� > 0 Then
        ʹ����� = "����"
    Else
        ʹ����� = "δ��"
    End If
    
End Property

Public Sub DeleteBill(ByVal intType As Integer)
    'inttype  : 0 ִ��ɾ��������1-ִ��UPdateʱ��ɾ������
    Dim strNO As String, strSQL As String
    On Error GoTo errHandle
    strNO = Me.NO
    
    If Me.ʹ����� <> "δ��" And intType = 0 Then
        If MsgBox("���ŵ����е�ҩƷ�Ѿ�ʹ�ã����ɾ�������ŵ�����ͬʹ�ü�¼��һͬ����������ָܻ���" & vbNewLine & " ��ȷ�ϣ��Ƿ�Ҫִ��ɾ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    ElseIf Me.ʹ����� <> "δ��" And intType = 1 Then
        MsgBox "���ŵ����е�ҩƷ�Ѿ�ʹ�ã�����ɾ����" & vbNewLine & " ��ȷ�ϣ��Ƿ�Ҫִ��ɾ��������", vbQuestion, gstrSysName
        Exit Sub
    ElseIf Me.ʹ����� = "δ��" And intType = 0 Then
        If MsgBox("ɾ�����ŵ��ݺ����ŵ��ݵ����ݽ���Զ��ʧ�����ָܻ���" & vbNewLine & " ��ȷ�ϣ��Ƿ�Ҫִ��ɾ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
     
    strSQL = "Zl_�ݴ�ҩƷ��¼_Delete('" & strNO & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "DeleteBill")
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub InsertUseBill(ByVal strBillKey As String, ByVal dblUse As Double, ByVal dateUse As Date)

    Dim objBIll As MediBill, strInsSQL As String
    
    On Error GoTo errHandle
    Set objBIll = Me.BillItem(strBillKey)
    If Not objBIll Is Nothing Then
        With objBIll
            If (.���� - .��������) >= dblUse And dblUse > 0 Then
                .ʹ��״̬ = 1
                strInsSQL = "ZL_�ݴ�ҩƷ��¼_Insert('" & mvarNO & "'," & .��� & "," & _
                         mvar����ID & "," & IIf(.ҽ��ID = 0, "Null", .ҽ��ID) & "," & IIf(.���ͺ� = 0, "Null", .���ͺ�) & "," & .ҩƷID & ",'" & _
                         .ҩƷ���� & "','" & .��� & "'," & .ִ�з��� & "," & _
                         .ʹ��״̬ & ",'" & mvarժҪ & "'," & -1 & ",'" & _
                         .���㵥λ & "'," & .���� & "," & dblUse & "," & _
                         .���� & "," & dblUse * .���� & ",'" & UserInfo.���� & "'," & _
                         mvar����ID & ",to_date('" & Format(dateUse, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'))"
                Call zlDatabase.ExecuteProcedure(strInsSQL, "MediMaster.InsertUseBill")
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub UndoUse(ByVal strBillKey As String)
    Dim objBIll As MediBill, strSQL As String
    On Error GoTo errHandle
    Set objBIll = Me.BillItem(strBillKey)
    If Not objBIll Is Nothing Then
        With objBIll
            If .ҽ��ID = 0 Then
                strSQL = "ZL_�ݴ�ҩƷ��¼_UndoUse('" & mvarNO & "'," & .��� & "," & .���ϵ�� & ",To_date('" & Format(.�Ǽ�ʱ��, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'))"
                Call zlDatabase.ExecuteProcedure(strSQL, "MediMaster.UndoUse")
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub