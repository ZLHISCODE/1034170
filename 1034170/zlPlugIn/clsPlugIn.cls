VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'---------------------------------------------------------------------------
'���ܣ���ҹ��ܽӿڲ���������һЩ����վ������ҵ��ù��ܣ���������Ӧ��ҵ�����
'�����ˣ����
'����ʱ�䣺2010-07-22
'˵����
'  2010-07-23:���ɵ�������/סԺҽ������ģ����
'  2010-11-25:�����˶��Զ����ù��ܵ�֧��,�Լ���ҽ��վǰ��ID�Ĵ���
'---------------------------------------------------------------------------

Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int���� As Integer = -1)
'���ܣ��ӿڳ�ʼ��
'������cnOracle=�����
'      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      �������� lngModual=1070����int���ϲ���
    Set gcnOracle = cnOracle
    
    Call HandlePlugIn(1, lngSys, lngModual, cnOracle, int����)
    
End Sub

Public Function GetFuncNames(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int���� As Integer = -1, Optional ByVal strReserve As String) As String
'���ܣ���ȡ��ҹ����б����ýӿ�һ���ڳ�ʼ��֮�����һ�Σ�����ʱ����Ƶ�ȿ�����ͬ�����������ͬ
'������lngSys,lngModual=��ǰ���ýӿڵ��ϼ�ϵͳ�ż�ģ��ţ���������ʱ lngModual=1070����int���ϲ���
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      strReserve=��������,������չʹ��
'      ҽ��վ�еĵ���ʱ,����Ϊ��ǰִ�е�ǰ��ҽ��ID
'���أ���ҹ����б�����ʽΪ"������1,������2,������3,..."
'      ���ĳ��������Ҫ�ڲ����л�(��ˢ��)ʱ�Զ����ã����ڹ�����ǰ�����"Auto:"ǰ׺�����ֹ���ֻ������һ����

'����˵�������������Ա����ģ�����ʱֻ����lngSys��lngModual���������е�lngModual��0
'          �Զ����ù�������Ա����ģ�����ʱ�����Զ�����

    Dim strTmp As String
    
    Call HandlePlugIn(2, lngSys, lngModual, , int����, strReserve, strTmp)
    
    GetFuncNames = "������1,������2,������3,������4,Auto:������5" & IIf(strTmp = "", "", "," & strTmp)
End Function

Public Sub ExecuteFunc(ByVal lngSys As Long, ByVal lngModual As Long, _
    ByVal strFuncName As String, ByVal lngPatiID As Long, ByVal varRecId As Variant, _
    ByVal varKeyId As Variant, Optional ByVal strReserve As String, Optional ByVal int���� As Integer = -1)
'���ܣ�ִ��ĳ����ҹ���
'������lngSys,lngModual=��ǰ���ýӿڵ��ϼ�ϵͳ�ż�ģ��ţ���������ʱ lngModual��1070����int���ϲ���
'      strFuncName=Ҫִ�еĹ��������ù�����֮ǰ��GetFuncNames��������
'      lngPatiID=��ǰ����ID
'      varRecId=���ֻ����ַ����������ﲡ�ˣ�Ϊ��ǰ�Һŵ��Ż��߹Һ�ID����סԺ���ˣ�Ϊ��ǰסԺ��ҳID
'      varKeyId=���ֻ����ַ�������ǰ�Ĺؼ�ҵ������Ψһ��ʶID����ҽ��ID
'           ��������ʱ varKeyId=���Ӳ�����¼.id���������༭δ����ǰ��ֵ=0,�������޸ġ���ʱִ����ֵ��
'           �ٴ�·�� varKeyId ����·��id
'      strReserve=��������,������չʹ�� ���ǲ�������ʱ strReserve�� �����ļ��б�.id
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)

'����˵�������������Ա����ģ�����ʱֻ����lngSys��lngModual��strFuncName��lngPatiID��varRecId��varKeyId��
'          lngPatiID������Աid�����lngPatiID��0��ʾ������δѡ����Ա�����е�varRecId��varKeyId��ֵ���Ǵ���0

    'ע�⣺�ڲ��÷������󣬹��ܵ�ִ��ʱ��Ӧ���ȶԹ����������жϡ�
    '�ȼ�鹦���Ƿ��ڱ������zlPlugIn���У�Ȼ����ȥִ����չ���
    
    If InStr("������1,������2,������3,������4,������5", strFuncName) > 0 Then
        MsgBox strFuncName & "ִ�гɹ���"
    Else
        Call HandlePlugIn(3, lngSys, lngModual, , int����, strReserve, strFuncName, lngPatiID, varRecId, varKeyId)
    End If
End Sub

Public Function QueryPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, _
    ByVal str���� As String, ByRef rsInfo As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����Ϣ����������Ϣ��Ϊ��̬��¼�����ڽӿ���Ԥ��д�ù����������ֶ�ģ�壻��������ֶ���"����ID"<>0�����ʾ�ӿڷ���ǰ���½��������Һų���ֱ����ȡ����ID��Ӧ�ĵ�����Ϣ������ֱ�Ӷ�ȡ�����ֶΣ�
    '���:lngModule-���õ�ģ���
    '����:rsInfo-���صĲ�����Ϣ��(ֱ�Ӹ��¸ñ�,�ֶ�Ϊ:����ID,����,�Ա�,����,��������,�����ص�,����֤��,����֤��,����,ְҵ,��ͥ��ַ,��ͥ�绰,��ͥ�ʱ�,������λ,��λ�ʱ�,ҽ����,ҽ�Ƹ��ʽ,�ѱ�,����,����,����״��,����)
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:33:19
    '˵��:׼�������������ϵͳ�ӿ�(��ȡ������Ϣ)
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '��:
    '    With rsInfo
    '        .AddNew
    '            !����ID = Null
    '            !���� = "��������" & Int(Rnd * 10)
    '            !�Ա� = "Ů"
    '            !���� = "44"
    '            !�������� = Now - 3000
    '            !�����ص� = "�����б����������Ŵ�1��"
    '            !����֤�� = "80223197405197890"
    '            !����֤�� = Null
    '            !���� = "����"
    '            !ְҵ = "IT"
    '            !��ͥ��ַ = "�山����ҵ԰��"
    '            !��ͥ�绰 = "67136410"
    '            !��ͥ�ʱ� = "5123"
    '            !������λ = "������Ϣ��ҵ"
    '            !��λ�ʱ� = "423"
    '            !ҽ���� = "42143Yb"
    '            !ҽ�Ƹ��ʽ = "��������"
    '            !�ѱ� = "��ͨ"
    '            !���� = "�й�"
    '            !���� = "����"
    '            !����״�� = "�ѻ�"
    '            !���� = "�山��"
    '        .Update
    '    End With
    QueryPatiInfo = True
End Function
 Public Function CommitPatiInfo(ByVal str���� As String, ByRef rsInfo As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ύ������Ϣ
    '���:str����-����
    '        rsInfo-���صĲ�����Ϣ��(ֱ�Ӹ��¸ñ�,�ֶ�Ϊ:����ID,����,�Ա�,����,��������,�����ص�,����֤��,����֤��,����,ְҵ,��ͥ��ַ,��ͥ�绰,��ͥ�ʱ�,������λ,��λ�ʱ�,ҽ����,ҽ�Ƹ��ʽ,�ѱ�,����,����,����״��,����)
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:33:19
    '˵��:׼�������������ϵͳ�ӿ�(�ύ������Ϣ)
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    CommitPatiInfo = True
End Function
Public Sub Terminate(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int���� As Integer = -1)
'���ܣ��ӿ���ֹ
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
    Call HandlePlugIn(4, lngSys, lngModual, , int����)
End Sub

Public Function ClinicReceive(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long) As Boolean
'���ܣ����˽�����������
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      lngPatientID=����ID,lngClinicID=����ID���Һż�¼ID��

    ClinicReceive = True
End Function

Public Function ClinicFinish(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long) As Boolean
'���ܣ�������ɽ�������
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      lngPatientID=����ID,lngClinicID=����ID���Һż�¼ID��

    ClinicFinish = True
End Function

Public Function DiagnosisEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lngԭ���id As Long, Optional ByVal int���� As Integer = -1) As Boolean
'���ܣ����ﲡ����������Ҫ��Ϻ����
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      lngPatientID=����ID,lngClinicID=����ID���Һż�¼ID��
'      lngDiagID=���ID�򼲲�ID(����ϵͳ��������),strDiagDesc=�������
'      lngԭ���id=0 ��ʾ�����ӵ���ϣ� ��Ϊ0��ʾ�޸���ϣ�lngԭ���id��ֵ�����޸�ǰ�� ���ID�򼲲�ID
'      int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
      
    DiagnosisEnter = True
End Function

Public Function DiagnosisOtherEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lngԭ���id As Long, Optional ByVal int���� As Integer = -1) As Boolean
'���ܣ����ﲡ���������Ҫ��Ϻ����
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      lngPatientID=����ID,lngClinicID=����ID���Һż�¼ID��
'      lngDiagID=���ID�򼲲�ID(����ϵͳ��������),strDiagDesc=�������
'      lngԭ���id=0 ��ʾ�����ӵ���ϣ� ��Ϊ0��ʾ�޸���ϣ�lngԭ���id��ֵ�����޸�ǰ�� ���ID�򼲲�ID
'      int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
    
    DiagnosisOtherEnter = True
End Function

Public Function DiagnosisDeleted(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional ByVal int���� As Integer = -1) As Boolean
'���ܣ����ﲡ����ɾ����Ϻ����
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      lngPatientID=����ID,lngClinicID=����ID���Һż�¼ID��
'      lngDiagID=���ID�򼲲�ID(����ϵͳ��������),strDiagDesc=�������
'      int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
    
    DiagnosisDeleted = True
End Function

Public Function AdviceEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional ByVal int���� As Integer = -1) As Boolean
'���ܣ�ÿ����һ��ҽ����Ŀʱ����
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���(1252=����ҽ���´�,1253=סԺҽ���´�)
'      lngPatientID=����ID,lngClinicID=����ID���Һż�¼ID��/��ҳID(סԺ����ʱ)
'      rsAdvice=��ǰ¼���һ����һ��ҽ����¼����ҩ�䷽�����׷�������������ҽ��������ʱ�����ã�
'      �ֶ�Ϊ��������ҽ����¼���Ĵ󲿷��ֶΣ�������ID, ���id, ���, ������Դ��1=����,2=סԺ), ����id, �Һŵ�/��ҳID,
'                ����, �Ա�, ����, Ӥ��, ҽ��״̬, ҽ����Ч, �������, ������Ŀid, �շ�ϸĿid, ����, ��������, �ܸ�����, ҽ������, ҽ������, �걾��λ,
'                ��鷽�� , ִ�б��, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��, �����λ, ִ��ʱ�䷽��, �Ƽ�����, ִ�п���id, ִ������, ������־, �ɷ����,
'                ��ʼִ��ʱ��,ִ����ֹʱ��(סԺ���ò��У�û��ʱΪcdate(0)), ���˿���id, ��������id, ����ҽ��, ����ʱ��, �Һŵ�, ǰ��id, ժҪ
'                ,����֮�⣬�������ʱ��������ҽ����Ӧ����ϵļ���ID
'      int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)

    AdviceEnter = True
End Function

Public Function AdviceSave(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional int���� As Integer = -1) As Boolean
'���ܣ�����ҽ��ʱ(����ʼǰ)����,����Falseʱ������ҽ��
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���(1252=����ҽ���´�,1253=סԺҽ���´�)
'      lngPatientID=����ID,lngClinicID=����ID���Һż�¼ID��/��ҳID(סԺ����ʱ)
'      rsAdvice=��ǰ���˱����¿�������ҽ����¼���ֶ�ͬAdviceEnter�ӿڵĸò���,����֮�⣬�������ʱ��������ҽ����Ӧ����ϵļ���ID
'      int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)

    AdviceSave = True
End Function


Public Function AdviceSend(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngSendID As Long) As Boolean
'���ܣ�����ҽ����ɺ����
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���(1252=����ҽ���´�(����û�С����͡�ģ��),1254=סԺҽ������)
'      lngPatientID=����ID,lngClinicID=����ID���Һż�¼ID��/��ҳID(סԺ����ʱ)
'      lngSendID=���ͺ�(����ҽ������.id)

    AdviceSend = True
End Function

Public Function AdviceRowChange(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional int���� As Integer = -1) As Boolean
'���ܣ�ҽ���༭ʱ�л�ҽ����ʱ����
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���(1252=����ҽ���´�,1253=סԺҽ���´�)
'      lngPatientID=����ID,lngClinicID=����ID���Һż�¼ID��/��ҳID(סԺ����ʱ)
'      rsAdvice=��ǰ���˱����¿�������ҽ����¼���ֶ�ͬAdviceEnter�ӿڵĸò���,����֮�⣬�������ʱ��������ҽ����Ӧ����ϵļ���ID
'      int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)

    AdviceRowChange = True
End Function

Public Function AdviceDeletBefor(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int���� As Integer = -1) As Boolean
'���ܣ�ҽ��ɾ��ǰ���ã�����False��ȡ��ɾ��ҽ����
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���(1252=����ҽ���´�,1253=סԺҽ���´�)
'      lngPatientID=����ID,lngClinicID=����ID���Һż�¼ID��/��ҳID(סԺ����ʱ)
'      lngAdviceID=����ҽ����¼.id
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)

    AdviceDeletBefor = True
End Function

Public Function AdviceDeleted(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int���� As Integer = -1) As Boolean
'���ܣ�ҽ��ɾ�������
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���(1252=����ҽ���´�,1253=סԺҽ���´�)
'      lngPatientID=����ID,lngClinicID=����ID���Һż�¼ID��/��ҳID(סԺ����ʱ)
'      lngAdviceID=����ҽ����¼.id
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)

    AdviceDeleted = True
End Function


Public Function AdviceRevoked(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int���� As Integer = -1) As Boolean
'���ܣ�ҽ�����Ϻ����
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���(1252=����ҽ���´�,1253=סԺҽ���´�)
'      lngPatientID=����ID,lngClinicID=����ID���Һż�¼ID��/��ҳID(סԺ����ʱ)
'      lngAdviceID=����ҽ����¼.id
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)

    AdviceRevoked = True
End Function

Public Function AddEMRBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant) As Boolean
'���ܣ��������Ӳ���ʱ�����
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���(���ﲡ�� 1250,סԺ���� 1251,�������� 1255)
'      lngPatientID=����ID,lngClinicID=����ID/
'      EMRFileID=�����ļ��б�.id
'      arrInput�����飺��չ����������չʹ�ã�ͨ���±��һֱ����ʱ���ܻᱨ���ִ���9���±�Խ��/13�����Ͳ�ƥ�䣺����ʾ����û�д���
'           ���Է��ʲ���ǰӦ���жϣ�If TypeName(arrInput(i)) <> "Error" Then ��������������ʹ������������������Ϊ�������û�д���
'           arrInput Ŀǰδ�����κβ���
    AddEMRBefore = True
End Function

Public Function SignEMRBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant) As Boolean
'���ܣ����Ӳ���ǩ��ǰ���ã�����δ����ʱ
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ��ţ�lngModual�̶�����1070
'      lngPatientID=����ID,lngClinicID=����ID/��ҳID
'      EMRFileID=���Ӳ�����¼.id
'      arrInput�����飺��չ����������չʹ�ã�ͨ���±��һֱ����ʱ���ܻᱨ���ִ���9���±�Խ��/13�����Ͳ�ƥ�䣺����ʾ����û�д���
'           ���Է��ʲ���ǰӦ���жϣ�If TypeName(arrInput(i)) <> "Error" Then ��������������ʹ������������������Ϊ�������û�д���
'           arrInput Ŀǰδ�����κβ���
'���أ�trueʱ����������ǩ�������򲻵���ǩ�����ڣ������ڲ�������ʾ��
    SignEMRBefore = True
End Function

Public Sub SignEMRAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant)
'���ܣ���ǩ�����ڹر����ĵ����ݱ�������
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ��ţ�lngModual�̶�����1070
'      lngPatientID=����ID,lngClinicID=����ID/��ҳID
'      EMRFileID=���Ӳ�����¼.id  �������༭δ����ǰ��ֵ=0�������ǩ�����޸�ʱǩ������ǩ��ֵ��
'      arrInput�����飺��չ����������չʹ�ã�ͨ���±��һֱ����ʱ���ܻᱨ���ִ���9���±�Խ��/13�����Ͳ�ƥ�䣺����ʾ����û�д���
'           ���Է��ʲ���ǰӦ���жϣ�If TypeName(arrInput(i)) <> "Error" Then ��������������ʹ������������������Ϊ�������û�д���
'           arrInput(0)��String��ǩ���˵�����

End Sub

Public Sub LogInAfter()
'���ܣ���½����̨֮��

End Sub

Public Sub LogOutBefore()
'���ܣ��رպ���̨����ע����ǰ
    
End Sub

Public Function PathImportBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng·��ID As Long, Optional ByVal lng�汾��ID As Long, Optional ByVal int���� As Integer = -1, _
    Optional ByVal lngDiagnosisType As Long, Optional ByVal lng����ID As Long, Optional ByVal lng���ID As Long, _
    Optional ByVal strReserve As String) As Boolean
'���ܣ�·������ǰ
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���(1252=����ҽ���´�,1253=סԺҽ���´�)
'      lngPatientID=����ID,lngClinicID=��ҳID
'      lng·��ID=�ٴ�·��Ŀ¼.id
'      lng�汾��ID=����İ汾��
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      lngDiagnosisType ����������� '�������:1-��ҽ�������;2-��ҽ��Ժ���;11-��ҽ�������;12-��ҽ��Ժ���
'      lng����ID=���뼲������������Ŀ¼.ID����lng���ID=�������(�������Ŀ¼.ID)
'      strReserve=��չ����

    PathImportBefore = True

End Function

Public Function PathImportAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng·��ID As Long, Optional ByVal lng�汾��ID As Long, Optional ByVal int���� As Integer = -1, _
    Optional ByVal lngDiagnosisType As Long, Optional ByVal lng����ID As Long, Optional ByVal lng���ID As Long, _
    Optional ByVal strReserve As String) As Boolean
'���ܣ�·������ǰ
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���(1252=����ҽ���´�,1253=סԺҽ���´�)
'      lngPatientID=����ID,lngClinicID=��ҳID
'      lng·��ID=�ٴ�·��Ŀ¼.id
'      lng�汾��ID=����İ汾��
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      strReserve=��չ����

    PathImportAfter = True

End Function

Public Function PathEvaluateBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng·����¼ID As Long, Optional ByVal lng�׶�ID As Long, Optional ByVal int���� As Integer = -1, Optional ByVal strReserve As String) As Boolean
'���ܣ�����·��������
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���1256=�ٴ�·��Ӧ��
'      lngPatientID=����ID,lngClinicID=��ҳID
'      lng·����¼ID=�����ٴ�·��.id
'      lng�׶�ID=��ǰ�����Ľ׶�ID
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      strReserve=��չ����

   PathEvaluateBefore = True

End Function

Public Function PathEvaluateAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng·����¼ID As Long, Optional ByVal lng�׶�ID As Long, Optional ByVal int���� As Integer = -1, Optional ByVal strReserve As String) As Boolean
'���ܣ�����·��������
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���1256=�ٴ�·��Ӧ��
'      lngPatientID=����ID,lngClinicID=��ҳID
'      lng·����¼ID=�����ٴ�·��.id
'      lng�׶�ID=��ǰ�����Ľ׶�ID
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      strReserve=��չ����

   PathEvaluateAfter = True

End Function

Public Function DiagnosisEnterIn(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lngԭ���id As Long, Optional ByVal int���� As Integer = -1, _
    Optional ByVal int������� As Integer, Optional ByVal int��ϴ��� As Integer, Optional ByVal strReserve As String) As Boolean
'���ܣ�סԺ������������Ϻ����
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      lngPatientID=����ID,lngClinicID=��ҳID
'      lngDiagID=���ID�򼲲�ID(����ϵͳ��������),strDiagDesc=�������
'      lngԭ���id=0 ��ʾ�����ӵ���ϣ� ��Ϊ0��ʾ�޸���ϣ�lngԭ���id��ֵ�����޸�ǰ�����ID�򼲲�ID
'      int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      int�������=1-��ҽ�������;2-��ҽ��Ժ���;3-��Ժ���(�������);5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���;10-����֢
'      int��ϴ���=��������µ�˳���
'      strReserve=��չ����
      
    DiagnosisEnterIn = True

End Function

Public Function GetFormHwnd(Optional ByVal strReserve As String) As Long
'���ܣ�������Ϣ���� ģ���е��� ��ȡҪǶ��Ĵ�������
'������
'      strReserve=��չ����
    GetFormHwnd = 0
End Function

Public Function PatiInfoSaveBefore(ByVal lng����ID As Long, Optional ByVal strReserve As String) As Boolean
'���ܣ�������Ϣ���� ģ���е��� ������Ϣ�����б��没����Ϣǰ�����ڿ���һЩ������ҵ�񣬷���False��ֹ���������
'������lng����ID   ����ID
'      strReserve=��չ����
    PatiInfoSaveBefore = True
End Function

Public Function PatiInfoSaveAfter(ByVal lng����ID As Long, Optional ByVal strReserve As String) As Boolean
'���ܣ�������Ϣ���� ģ���е��� ������Ϣ�����б��没����Ϣ�������ύ����ã�������������д������ݱ���
'������lng����ID   ����ID
'      strReserve=��չ����
    PatiInfoSaveAfter = True
End Function