VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBillPrint"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'˵����
'    �����������ݷ��ʲ���ms ado 6.0:msado15.dll����ͨ��ZLHIS������װ����װ���򵥶���װMDAC_TYP_V2.6
'    ��ģ����Ϊͨ�õ��ⲿ�ӿڵ��ã�ÿ��������Ʊ�ݴ�ӡ������װ��һ��������ģ���ļ���
'    mdlBillPrint_BJ=����Ʊ�ݴ�ӡ�ӿ�
'ע�⣺
'   ����������ӿ���C��C#�����Ա�д�ģ���������Ϊ�ֽ�������ʱ��VB����Ҫ����Ϊ�����ַ����ͣ������Ƚ��г�ʼ��
'   ���磺dim strError As String * 255      strError = String(255, " ")

Public Function zlInitialize(ByRef cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal lngModul As Long, _
    ByVal strUserCode As String, ByVal strUserName As String) As Boolean
'���ܣ���ʼ���ӿ�,��HIS����ģ��ʱ����(���磺�����շѹ�������)�����ڱ��ӿ�����ɵ�����Ʊ�ݴ�ӡ�ӿڵĳ�ʼ������¼�������ӵȵ���

'������cnMain=���ݿ����Ӷ���
'      strUserCode=��ǰ����Ա���
'      strUserName=��ǰ����Ա����
'      lngSys=��ǰ����ϵͳ��ţ�100=ZLHIS��׼��
'      lngModul=��ǰ����ģ��ţ�1121=�����շ�,1137=סԺ����
'
'���أ���ʼ���ɹ�/ʧ��
        
    Set gcnOracle = cnMain
    glngSys = lngSys
    glngModul = lngModul
    gstrUserCode = strUserCode
    gstrUserName = strUserName
    
    zlInitialize = Init
End Function

Public Function zlTerminate() As Boolean
'���ܣ���ֹ�ӿڣ���HIS�˳�ģ��ʱ����(���磺�˳��շѹ�������)�����ڱ��ӿ�����ɵ�����Ʊ�ݴ�ӡ�ӿڵ���Դ�ͷš��Ͽ����ӵȵ���
'���أ�ִ�гɹ�/ʧ��
    
    zlTerminate = Term
End Function

Public Function zlConfigure() As Boolean
'���ܣ���������,��HIS�շѻ����ģ���������"(�ļ�/��������)��ִ��"Ʊ�ݴ�ӡ����"ʱ���ã����ڱ��ӿ�����ɵ�����Ʊ�ݴ�ӡ�ӿڵĲ������á����ø��ĵȵ��á�
'���أ�ִ�гɹ�/ʧ��
    
    zlConfigure = SYSConfigure
End Function

Public Function zlDiscardBill(ByVal lng����ID As Long, ByVal lngƱ�� As Long, ByVal strƱ��ǰ׺ As String, _
    ByVal str��ʼƱ�� As String, ByVal str����Ʊ�� As String, ByVal DateAdd As Date, ByVal str������ As String) As Boolean
'���ܣ�Ʊ�ݱ��𣬱�����ɺ���ñ��ӿ�
'������lng����ID = Ʊ�����ü�¼.ID������������ID
'      lngƱ��   = 1-�շ�Ʊ��,2-Ԥ��Ʊ��,3-����Ʊ��,4-�Һ�Ʊ��,5-���￨
'      DateAdd   = ����ʱ��
'      str������ = ����������

'���أ�ִ�гɹ�/ʧ��
    
    zlDiscardBill = DiscardBill(lng����ID, lngƱ��, strƱ��ǰ׺, str��ʼƱ��, str����Ʊ��, DateAdd, str������)
End Function

Public Function zlPrintBill(ByVal strNOs As String, ByVal lngBalanceId As Long) As Boolean
'���ܣ������շѻ�סԺ������ɺ����Ʊ�ݴ�ӡ
'      �����޸�һ�ŵ��ݺţ��ȵ���Ʊ�����Ͻӿ�zlEraseBill���ٵ��ñ��ӿڡ�
'������strNOs=�����շѣ��Զ��ŷָ��Ĵ����ŵĶ�����ݺ�(һ�δ�ӡ���Ż���ŵ���):'F0000001','F0000002',...
'      lngBalanceId=סԺ���ʣ����ʵ�ID
'���أ�ִ�гɹ�/ʧ��
    If glngModul = FOutBillPrint Then
        '�����շ�
        zlPrintBill = PrintBillOut(strNOs)
    Else
        'סԺ����
        zlPrintBill = PrintBillIn(lngBalanceId)
    End If
End Function


Public Function zlRePrintBill(ByVal strNOs As String, ByVal lngBalanceId As Long, ByVal strInvoice As String) As Boolean
'���ܣ����´�ӡƱ�ݣ�����ֱ��ʹ���ش��ܣ��Լ����ﲿ���˷�ʱ�ش����������
'       ���ﲿ���˷��ش�ʱ���ȵ���zlEraseBill�ӿ�����Ʊ�ݣ��ٵ��ñ��ӿڡ�
'       ����ȫ���˷�ʱ���ջ�Ʊ�ݣ������ñ��ӿڣ�ֻ�������Ͻӿڡ�

'������strNOs=�����շѣ��Զ��ŷָ��Ĵ����ŵĶ�����ݺ�(һ�δ�ӡ���Ż���ŵ���):'F0000001','F0000002',...
'      lngBalanceId=סԺ���ʣ����ʵ�ID
'      strInvoice=�����ش�ʹ�õ���ʼƱ�ݺ�
'���أ�ִ�гɹ�/ʧ��
    If glngModul = FOutBillPrint Then
        '�����շ�
        zlRePrintBill = RePrintBillOut(strNOs, strInvoice)
    Else
        'סԺ����
        zlRePrintBill = RePrintBillIn(lngBalanceId, strInvoice)
    End If
End Function


Public Function zlEraseBill(ByVal strNOs As String, ByVal lngBalanceId As Long) As Boolean
'���ܣ������Ѵ�ӡƱ�ݣ�����ȫ�˺�סԺ��������ʱ����
'      �����޸ĵ����ش�ǰ�Ͳ����˷��ش�ǰ�ȵ��ñ��ӿڡ�
'������strNOs=�����շѣ��Զ��ŷָ��Ĵ����ŵĶ�����ݺ�(һ�δ�ӡ���Ż���ŵ���):'F0000001','F0000002',...
'      lngBalanceId=סԺ���ʣ����ʵ�ID
'���أ�ִ�гɹ�/ʧ��
    If glngModul = FOutBillPrint Then
        '�����շ�
        zlEraseBill = EraseBillOut(strNOs)
    Else
        'סԺ����
        zlEraseBill = EraseBillIn(lngBalanceId)
    End If
End Function

Public Function zlBillInCheckValied(ByVal bytOperType As Byte, ByVal bytBillType As Byte, ByVal strUseType As String, ByVal strFactBatchNum As String, _
                                    ByVal strStartFactNo As String, ByVal strEndFactNo As String, ByRef strExpended As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ʊ�ݵĺϷ���
    '���:bytOperType=�������ͣ�1-�������Ʊ�ݣ�2-�޸����Ʊ�ݣ�3-ɾ�����Ʊ��
    '    :bytBillType=Ʊ�����ͣ� 1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
    '    :strUseType=ʹ�����ͣ� 1,4: Ʊ��ʹ�����.����;2Ԥ��:1-����Ԥ��;2-סԺԤ��;5:�洢����ҽ�ƿ����.ID��ɾ��ʱ��Ч
    '    :strFactBatchNum=Ʊ������'���޸ĺ�ɾ��ʱ��Ч
    '    :strStartFactNo=��ʼ��Ʊ�ţ�ɾ��������Ч
    '    :strEndFactNo=��ֹ��Ʊ�ţ�ɾ��������Ч
    '    :strExpended=��չ��������ʱδʹ��
    '����:ִ�гɹ�/ʧ��
    '����:���ϴ�
    '����:2016/11/22 11:09:34
    '˵����
    '   �����Ǽ�ʱ����ȷ��ʱ���øýӿڡ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlBillInCheckValied = True
End Function

Public Function zlBillIn(ByVal bytOperType As Byte, ByVal bytBillType As Byte, ByVal strUseType As String, ByVal strFactBatchNum As String, _
                        ByRef strExpended As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��Ʊ�����
    '���:bytOperType=�������ͣ�1-�������Ʊ�ݣ�2-�޸����Ʊ�ݣ�3-ɾ�����Ʊ��
    '    :bytBillType=Ʊ�����ͣ�1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
    '    :strUseType=ʹ�����ͣ� 1,4: Ʊ��ʹ�����.����;2Ԥ��:1-����Ԥ��;2-סԺԤ��;5:�洢����ҽ�ƿ����.ID
    '    :strFactBatchNum=Ʊ������
    '    :strExpended=��չ��������ʱδʹ��
    '����:ִ�гɹ�/ʧ��
    '����:���ϴ�
    '����:2016/11/22 11:09:34
    '˵����
    '   �����Ǽ�ʱ����ȷ��ʱ���øýӿڡ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlBillIn = True
End Function
