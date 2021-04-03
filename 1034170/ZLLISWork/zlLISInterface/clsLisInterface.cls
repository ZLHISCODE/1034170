VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function gOpenDataBase(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '===================================================================================
    '����:                                  �����ݿ�����
    '����;
    'strServerName                          �����������ִ�
    'strUserName                            �û���
    'strUserPwd                             ����
    '����                                   True=���ӳɹ� False=����ʧ��
    '===================================================================================
    gOpenDataBase = OraDataOpen(strServerName, strUserName, strUserPwd)
End Function
Public Function gGetDeptPatiList(ByVal strDeptNo As String, ByRef strReturn As String, ByVal lngType As Long, ByVal strStartDate As String, ByVal strEndDate As String, ByRef ErrInfo As String) As Boolean
    '===================================================================================
    '����:                                  ��������ȡ���������嵥
    '����;
    'strDeptNo                              ���ұ���
    'strReturn                              ���������б�
    'lngType                                0-��һ�ε��ã�ȡ100�����룬>0ȡ���������룬ֱ�����ؿմ���
    'strStartDate                           ��ʼ����,��ʽ YYYY-MM-DD
    'strEndDate                             ��������,��ʽ YYYY-MM-DD
    'ErrInfo                                ��ʾ��Ϣ
    '����                                   True=�ɹ� False=ʧ��
    '===================================================================================
    gGetDeptPatiList = GetDeptPatiList(strDeptNo, strReturn, lngType, strStartDate, strEndDate, ErrInfo)

End Function
Public Function gGetApplication(strPatientID As String) As String
    '=========================================================================================
    '����:                              �õ��������뵥�ļ�¼��
    '����
    'strPatientID                       ����Ϊ���￨�Ż����롢��������ͷΪ����ID��������סԺ�š���*������š���.���Һŵ��š���/���շѵ��ݺš���\��������
    '����                               �����ִ�����:"ҽ��ID^����^�Ա�^����^������Դ^�����^סԺ��^������ұ���^�����������^����ҽ��^����ʱ��^ҽ�����^��������^�����ӡ����^ҽ�����^��ǰ����^���˿��ұ���^���˿�������^Ӥ��^���"
    '                                   Ϊ�ձ�ʾû���ҵ���¼
    '=========================================================================================
    gGetApplication = GetApplication(strPatientID)
End Function

Public Function gOraDataClose() As Boolean
    '==============================================================
    '����                               '�ر����ݿ�����
    '==============================================================
    gOraDataClose = OraDataClose
End Function

Public Function gInsertReport(lngID As Long, strReportPath As String, strErrInfo As String) As Boolean
    '===================================================================
    '����                               ���뱨�浽HIS
    '����
    'lngID                              ҽ��ID
    'strReportPath                      ����·��
    '===================================================================
    gInsertReport = InsertReport(lngID, strReportPath, strErrInfo)
End Function

Public Function gDeleteReport(lngID As Long) As Boolean
    '===================================================================
    '����                               ɾ������
    '����
    'lngID                              ҽ��ID
    '===================================================================
    gDeleteReport = DeleteReport(lngID)
End Function

Public Function gGetClinicItem(lngAdivce As Long) As String
    '===================================================================
    '����                               ȡ��Ҫ����������Ŀ����
    '����
    'lngAdivce                          ҽ��ID
    '����                               �ִ���ʽ:������ĿID^������Ŀ����^������Ŀ����^ִ�п��ұ���^ִ�п�������^����^���^�Ƿ��շ�
    '===================================================================
    gGetClinicItem = GetClinicItem(lngAdivce)
End Function

Public Function gGetItemList(lngClinicID As Long) As String
    '===================================================================
    '����                               ȡ��������Ŀ��ָ����ϸ
    '����
    'lngClinicID                        ������ĿID
    '����                               �ִ���ʽ:����^������^Ӣ����^����
    '===================================================================
    gGetItemList = GetItemList(lngClinicID)
End Function

Public Function gSetRegister(lngAdivce As Long, intTag As Integer) As Boolean
    '=====================================================================
    '����                               �걾���ջ�ȡ������
    '����
    'lngAdivce                          ҽ��ID
    'intTag                             1=���� 0=ȡ������
    '=====================================================================
    gSetRegister = SetRegister(lngAdivce, intTag)
End Function

Public Function gGetAllItem(Optional strFindItem As String) As String()
    '=====================================================================
    '����                               ȡ�����е�������Ŀ���������
    '����
    'strItem                            ��ѡ�����ұ����������ͬ��������Ŀ��Ŀ
    '����                               ���ҵ���������Ŀ����
    '=====================================================================
    gGetAllItem = GetAllItem(strFindItem)
End Function

Public Function gTestResults(ByVal lngID As Long, ByVal strTestName As String, ByVal strTestTime As String, ByVal strTestResults As String) As String
    '===================================================================
    '����                               ���ؼ����������ϵͳ
    '����
    'lngID                              ҽ��ID
    'strTestName                        ������
    'strTestTime                        ����ʱ�䣬��ʽ 2009-01-01 10:30:01
    'strTestResults                     ҽ��ID��Ӧ�ļ����������ԶԶ��ٸ�����ָ��һ��������ϸ��ʽ���£�
    '
    '                                     ������Ŀid;������1;��λ1;�����1��;�����־1|������Ŀid;������2;��λ2;����ο�2;�����־2......
    '
    '                                     ���У������־�� ��ƫ��,ƫ��,�쳣,�մ�����ѡ��һ�����ء�
    '����: �գ���ʾ���³ɹ����ǿգ���ʾ������Ϣ��
    '===================================================================
    gTestResults = UpdateTestResults(lngID, strTestName, strTestTime, strTestResults)
End Function

Public Function gzlLisRegister(ByVal lngDevID As Long, ByVal lngID As Long, ByVal strSampleNo As String, ByRef strErrInfo As String) As Boolean
    '===================================================================
    '����                               �������뵽zllis
    '����
    'lngDevID                           ����ID
    'lngID                              ҽ��ID
    'strSampleNo                        �걾��


    '===================================================================
    gzlLisRegister = zlLISRegister(lngDevID, lngID, strSampleNo, strErrInfo)
End Function

Public Function gzlLisUnRegister(ByVal lngID As Long, ByRef strErrInfo As String) As Boolean
    '===================================================================
    '����                               ȡ����zllis���Ѻ��յı걾
    '����
    'lngDevID                           ����ID
    'lngID                              ҽ��ID
    'strSampleNo                        �걾��

    '����: True����ʾ���³ɹ���false����ʾʧ�ܣ���Ϣ��strErrInfo�����С�
    '===================================================================
    gzlLisUnRegister = zlLisUnRegister(lngID, strErrInfo)
End Function

Public Function gZLLisInsterReport(ByVal lngID As Long, ByVal strItems As String, ByRef strErrInfo As String) As Boolean
    'д������˱���
    gZLLisInsterReport = ZLLisInsterReport(lngID, strItems, strErrInfo)
End Function

Public Function gzlLisUnAudit(ByVal lngID As Long, ByRef strErrInfo As String) As Boolean
    'ȡ�����󱨸�
    gzlLisUnAudit = zlLisUnAudit(lngID, strErrInfo)
End Function

Public Function gGetAllDevice(ByRef strErrInfo As String) As String
    'ȡ��������
    gGetAllDevice = GetAllDevice(strErrInfo)
End Function

Public Function gIncomeverify(ByVal lngID As Long, ByRef strErr As String) As Boolean
    '��˻��۵�
    gIncomeverify = Incomeverify(lngID, strErr)
End Function

Public Function gCriticalvalueNotice(ByVal lngID As Long, ByVal strNoticeTitle As String, ByVal strNotice As String) As Long
    'Σ��ֵ֪ͨ����
    gCriticalvalueNotice = CriticalvalueNotice(lngID, strNoticeTitle, strNotice)
End Function
Private Sub Class_Initialize()
    If gobjComLib Is Nothing Then Set gobjComLib = GetObject("", "zl9Comlib.clsComlib")
End Sub
