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
'���ܣ���չ�ӿڲ��ģ��
'���Ҫ������ʹ�� ��ʵ�� GetUserName ����
'---------------------------------------------------------------------------

Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int���� As Integer = -1)
'���ܣ��ӿڳ�ʼ��
'������cnOracle=�����
'      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
    
    Set gcnOracle = cnOracle
    
End Sub

Public Function GetFuncNames(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int���� As Integer = -1, Optional ByVal strReserve As String) As String
'���ܣ���ȡ��ҹ����б����ýӿ�һ���ڳ�ʼ��֮�����һ�Σ�����ʱ����Ƶ�ȿ�����ͬ�����������ͬ
'������lngSys,lngModual=��ǰ���ýӿڵ��ϼ�ϵͳ�ż�ģ���
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      strReserve=��������,������չʹ��
'        ҽ��վ�еĵ���ʱ,����Ϊ��ǰִ�е�ǰ��ҽ��ID
'���أ���ҹ����б�����ʽΪ"������1,������2,������3,..."
'      ���ĳ��������Ҫ�ڲ����л�(��ˢ��)ʱ�Զ����ã����ڹ�����ǰ�����"Auto:"ǰ׺�����ֹ���ֻ������һ����
    
    GetFuncNames = "������1,������2,������3,������4,������5"
End Function

Public Sub ExecuteFunc(ByVal lngSys As Long, ByVal lngModual As Long, _
    ByVal strFuncName As String, ByVal lngPatiID As Long, ByVal varRecId As Variant, _
    ByVal varKeyId As Variant, Optional ByVal strReserve As String, Optional ByVal int���� As Integer = -1)
'���ܣ�ִ��ĳ����ҹ���
'������lngSys,lngModual=��ǰ���ýӿڵ��ϼ�ϵͳ�ż�ģ���
'      strFuncName=Ҫִ�еĹ��������ù�����֮ǰ��GetFuncNames��������
'      lngPatiID=��ǰ����ID
'      varRecId=���ֻ����ַ����������ﲡ�ˣ�Ϊ��ǰ�Һŵ��Ż��߹Һ�ID����סԺ���ˣ�Ϊ��ǰסԺ��ҳID
'      varKeyId=���ֻ����ַ�������ǰ�Ĺؼ�ҵ������Ψһ��ʶID����ҽ��ID
'      strReserve=��������,������չʹ��
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
    
    MsgBox strFuncName & "ִ�гɹ���"
    
End Sub
 
Public Function GetUserName() As String
'���ܣ���ȡ��λ����
'˵�����û�����ʱ��ʾ������
'��  GetUserName = "������������ҽԺ��Ϣϵͳ"

    GetUserName = ""
End Function
 
Public Sub Terminate(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int���� As Integer = -1)
'���ܣ��ӿ���ֹ
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
     
End Sub
