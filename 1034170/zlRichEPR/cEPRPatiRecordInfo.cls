VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRPatiRecordInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'��¼��Ϣ
Public ID As Long                          '��¼ID
Public ������Դ As PatiFromEnum            '1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
Public ����ID As Long                      '
Public ��ҳID As Long                      'סԺ���˱�����д
Public Ӥ�� As Long                        '�����Ӥ������ʾʱ����������Ϊ**֮Ӥ
Public ����ID As Long                      '�������,��¼��д���β�����ҽ����ʿ���ڿ���
Public �������� As EPRDocTypeEnum          '��Ӧ�ڲ����ļ�������
Public �ļ�ID As Long                      '�����Ӧ�Ĳ����ļ��Ѿ�ɾ��,���ʽ��ȱʡ������
Public �������� As String                  '
Public ������ As String
Public ����ʱ�� As Date
Public ������ As String
Public ����ʱ�� As Date
Public �鵵�� As String
Public �鵵���� As Date
Public ���ʱ�� As Date                    '��ǩ��ʱ��д,ƽ�������µİ汾����ʱ����д,��ʾ���ڽ�����
Public ���汾 As Integer
Public ǩ������ As EPRSignLevelEnum        '1-��д;2-����ҽʦ����;3-����ҽʦ���ġ�סԺ��������Ĳ���ֻ����д������״̬
Public �༭��ʽ As Integer                  ' 0-ʹ��ȫ��ʽ���Ӳ����༭���༭��1-ʹ�ñ���ʽ�����༭���༭��2-���ﲡ���������(2����0)

'���䣺
Public ҽ��id As Long                      'ҽ��ID���������ⲡ����

Private mvar��дǩ�� As String
Private mvarҽ��ǩ�� As String
Private mvar����ǩ�� As String
Private mvar����ǩ�� As String

Public Property Get ��д����() As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    ��д���� = ""
    Err = 0: On Error GoTo errHand
    strSQL = "Select ���� From ���ű� Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "cEPRPatiRecordInfo", Me.����ID)
    If rsTemp.RecordCount > 0 Then ��д���� = Trim(rsTemp.Fields(0).Value)
    rsTemp.Close
    Set rsTemp = Nothing
    
    Exit Property
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0
End Property
Public Property Let ҽ��ǩ��(ByVal vData As String)
    mvarҽ��ǩ�� = vData
End Property

Public Property Get ҽ��ǩ��() As String
    ҽ��ǩ�� = mvarҽ��ǩ��
End Property
Public Property Let ����ǩ��(ByVal vData As String)
    mvar����ǩ�� = vData
End Property

Public Property Get ����ǩ��() As String
    ����ǩ�� = mvar����ǩ��
End Property
Public Property Let ����ǩ��(ByVal vData As String)
    mvar����ǩ�� = vData
End Property

Public Property Get ����ǩ��() As String
    ����ǩ�� = mvar����ǩ��
End Property
Public Property Let ��дǩ��(ByVal vData As String)
    mvar��дǩ�� = vData
End Property

Public Property Get ��дǩ��() As String

    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    If mvar��дǩ�� = "" Then
    
        Err = 0
        On Error GoTo errHand
        
        strSQL = "Select �����ı� From ���Ӳ������� Where �ļ�id = [1] And �������� = 8 And ��ʼ�� = 1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "cEPRPatiRecordInfo", Me.ID)
        If rsTemp.RecordCount > 0 Then mvar��дǩ�� = Trim(rsTemp.Fields(0).Value)
        rsTemp.Close
        Set rsTemp = Nothing

    End If
    
    ��дǩ�� = mvar��дǩ��
    

    Exit Property
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0

End Property