VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Seating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'��������ֵ�ľֲ�����
Private mvar��� As Integer '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar����� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar�Ա� As String '�ֲ�����
Private mvar��� As String '�ֲ�����
Private mvar��� As Integer '�ֲ�����
Private mvar״̬ As Integer '�ֲ�����
Private mvar�ּ� As Currency '�ֲ�����
Private mvar�շ�ϸĿID As Long '�ֲ�����
Private mvar�շ���Ŀ As String '�ֲ�����
Private mvar��ע As String '�ֲ�����
Private mvar����  As Long  '�ֲ�����
Private mvarKey As String

'---- ����ƶ��ٴ��޸�
Private mvar��������� As String    '
Private mvar��ʼʱ�� As String
Private mvar���     As String
Private mvar����    As String
Private mvar�Һŵ�  As String
Private mstrPatientKey As String

Public Property Let Key(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�շ�ϸĿID = 5
    mvarKey = vData
End Property


Public Property Get Key() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�շ�ϸĿID
    Key = mvarKey
End Property

'���˵�Key���������ID_�Һŵ��ţ��������ۣ�����ID_��ҳID
Public Property Let PatientKey(ByVal strVar As String)
    mstrPatientKey = strVar
End Property
Public Property Get PatientKey() As String
    PatientKey = mstrPatientKey
End Property

Public Property Let ��ע(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�շ�ϸĿID = 5
    mvar��ע = MidUni(vData, 1, 100)
End Property


Public Property Get ��ע() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�շ�ϸĿID
    ��ע = mvar��ע
End Property

Public Property Let �շ���Ŀ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�շ�ϸĿID = 5
    mvar�շ���Ŀ = vData
End Property


Public Property Get �շ���Ŀ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�շ�ϸĿID
    �շ���Ŀ = mvar�շ���Ŀ
End Property

Public Property Let �շ�ϸĿID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�շ�ϸĿID = 5
    mvar�շ�ϸĿID = vData
End Property


Public Property Get �շ�ϸĿID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�շ�ϸĿID
    �շ�ϸĿID = mvar�շ�ϸĿID
End Property



Public Property Let �ּ�(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ּ� = 5
    mvar�ּ� = vData
End Property


Public Property Get �ּ�() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�ּ�
    �ּ� = mvar�ּ�
End Property

Public Property Let ����(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    ���� = mvar����
End Property

Public Property Let ״̬(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.״̬ = 5
    mvar״̬ = vData
End Property


Public Property Get ״̬() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.״̬
    ״̬ = mvar״̬
End Property



Public Property Let ����(ByVal vData As String)
    mvar���� = vData
End Property


Public Property Get ����() As String
    ���� = mvar����
End Property
Public Property Let ���(ByVal vData As Integer)
    mvar��� = vData
End Property


Public Property Get ���() As Integer
    ��� = mvar���
End Property


Public Property Let ���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��� = vData
End Property


Public Property Get ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���
    ��� = mvar���
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

Public Property Let �Һŵ�(ByVal vData As String)
    mvar�Һŵ� = vData
End Property

Public Property Get �Һŵ�() As String
    �Һŵ� = mvar�Һŵ�
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

Public Property Let ���������(ByVal vData As String)
    mvar��������� = vData
End Property

Public Property Get ���������() As String
    ��������� = mvar���������
End Property

Public Property Let ��ʼʱ��(ByVal vData As String)
    mvar��ʼʱ�� = vData
End Property

Public Property Get ��ʼʱ��() As String
    ��ʼʱ�� = mvar��ʼʱ��
End Property

Public Property Let ���(ByVal vData As String)
    mvar��� = vData
End Property

Public Property Get ���() As String
    ��� = mvar���
End Property

Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar����ID = vData
End Property


Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����ID
    ����ID = mvar����ID
End Property

Public Property Let ���(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��� = vData
End Property


Public Property Get ���() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���
    ��� = mvar���
End Property

Public Function Update(����ID As Long, �շ�ϸĿID As Long, ״̬ As Integer, �շ���Ŀ As String, �ּ� As Currency, ��ע As String, ���� As Long, ������ As String) As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select A.����ID,B.���� From ��λ״����¼ A,������Ϣ B Where  A.����ID=B.����ID(+) And A.����ID=[1] And A.���=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Seating.Update", ����ID, CStr(mvar���))
    If rsTmp.EOF Then
        Update = "��λ�ѱ�ɾ����������ѡ����λ��"
        Exit Function
    End If
    
    If rsTmp!����ID <> 0 Then
        '��λ���ˣ����ܸ���
        Update = "�ѱ�" & rsTmp!���� & "(" & ����ID & ") ռ��"
    Else
        '������λ �� �����λ
       
        strSQL = "ZL_��λ״����¼_UPDATE(" & ����ID & ",'" & mvar��� & "'," & IIf(�շ�ϸĿID = 0, "Null", �շ�ϸĿID) & "," & ״̬ & ",'" & DelInvalidChar(��ע, "'") & "'," & ���� & ",'" & ������ & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "seating.Update")
        mvar�շ�ϸĿID = �շ�ϸĿID
        mvar�շ���Ŀ = �շ���Ŀ
        mvar�ּ� = �ּ�
        mvar��ע = ��ע
        mvar��������� = ������

    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
