VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTestDeug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'��������ֵ�ľֲ�����
'�ӿ���
Private Function zlInitComponents(ByVal intCallType As Integer, cnOracle As ADODB.Connection) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '���:intCallType Byte    HIS���� 0-  ������õ���  1-  סԺ���ʵ���
    '     cnOracle    ADODB.Connection    HIS���� �������Ӷ���
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlInitComponents = True
End Function
Public Function zlBrushCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal strPatiName As String, ByVal strSex As String, _
    ByVal strOld As String, ByVal dbl��� As Double, _
    Optional ByRef strCardNo As String, _
    Optional ByRef strPassWord As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ��֧�����,����
    '���:rsClassMoney:�շ����,���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-18 14:18:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlBrushCard = True
 End Function
Public Function zlCardDevSet(ByVal frmMain As Object, Optional lngCardCode As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:zlCardDevSet(�豸�������ýӿ�):��Ҫ��������ص��豸����ز���
    '���:frmMain Form    HIS���� ���õ�������
    '����:
    '����:true:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:18:38
    'HIS����˵��.
    '    1.  �������շѹ����Ĳ�������'�豸����' zlCardDevSet
    '    2.  ��סԺ���ʽ���Ĳ�������'�豸����' zlCardDevSet
    '---------------------------------------------------------------------------------------------------------------------------------------------
    frmSetDrvParaSet.ShowMe frmMain, lngCardCode
End Function

Public Function zlBrushCardSquare(ByVal intCallType As Integer, frmMain As Object, ByVal lng�ӿڱ�� As Long, ByVal rsFeeList As ADODB.Recordset, ByVal dbl������� As Double, ByRef rsSquare As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: zlBrushCardSquare (ˢ������ӿ�)
    '���:frmMain:HIS���� ���õ�������
    '     intCallType : HIS���� 0-  ������õ��� 1-  סԺ���ʵ���
    '     rsFeeList: HIS���� ���������൥��,�����е��ݵ���ϸ,�����סԺ���� , ���Ǳ��ν��ʵ�������ϸ
    '     dbl������� :  HIS���� ��ʾˢ�����ܳ����˽��
    '
    '����:rsSquare : �ӿڷ���    ���ؼ�¼��:�ӿڴ���սṹ(�ӿڷ�����ص�����) , �ṹ����:
    '                �ӿڱ�� , ���ѿ�ID, ���㷽ʽ, ������, ���ſ�����, ������ˮ��, ����ʱ��, ��ע
    '     rsSquare˵��:��Ҫ�ǽ��ͬһ����,ˢ���ſ����ѵ����.,�������ˢ���ſ� , ����ӿ����Ѿ�ˢ���Ŀ���Ϣ
    '����:true:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:18:38
    '˵��:
    '    1.  �������շѽ���ʱ,HIS�ڵ�"���㿨"ʱ,���ñ��ӿ�
    '    2.  ��סԺ���ʽ���ʱ,HIS�ڵ�"���㿨"ʱ,���ñ��ӿ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
End Function
Public Function zlSquareFee(ByVal intCallType As Integer, frmMain As Object, ByVal str����ID_IN As String, ByVal rsSquare As ADODB.Recordset) As Boolean
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: zlSquareFee (����ӿ�)
    '���:frmMain:HIS���� ���õ�������
    '     intCallType : HIS���� 0-  ������õ��� 1-  סԺ���ʵ���
    '     str����ID_IN: HIS���� ���ν��ʵĽ���ID��
    '     rsSquare :  ����Ӧˢ���Ľ���
    '����:
    '����:true:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:18:38
    '˵��:
    '    1. ��"�����շ�"�����"ȷ��"ʱ,���ñ��ӿ�
    '    2. ��"סԺ����"�����"ȷ��"ʱ,���ñ��ӿ�
    'ע:
    '  �˽ӿ���������HIS������ , ��˲����ڴ˽ӿڴ������û������Ĳ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    'strFile = strPath & "\BrushCard.txt"
End Function
Public Function zlCheckDelSquareValied(frmMain As Object, ByVal intCallType As Integer, ByVal str����ID_IN As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ִ���˷�ʱ,�����صĽӿڲ����Ƿ�����
    '���:
    '����:
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-31 16:39:47
    '˵��;
    '     ���˷�ʱ����Ҫ������صļ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
End Function
Public Function zlDelSquareFee(ByVal intCallType As Integer, frmMain As Object, ByVal str����ID_IN As String) As Boolean
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: zlSquareFee (����ӿ�)
    '���:frmMain:HIS���� ���õ�������
    '     intCallType : HIS���� 0-  ������õ��� 1-  סԺ���ʵ���
    '     str����ID_IN: HIS���� ���ν��ʵĽ���ID��
    '����:
    '����:true:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:18:38
    '˵��:
    '    1. "�����շѹ���"��"סԺ���ʹ���"������ʱ,���ô˽ӿ�
    'ע:
    '  �˽ӿ���������HIS������ , ��˲����ڴ˽ӿڴ������û������Ĳ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
End Function
Public Function zlSimulationBrushCard(ByVal frmMain As Object, ByVal Dblˢ����� As Double, _
    ByRef lng�ӿڱ�� As Long, ByRef strBlanceInfor As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ѡ��ָ��������
    '��Σ�frmMain HIS���� ���õ�������
    '         Dblˢ����� HIS���� ����Ԥ�������еĽ��
    '         Lng�ӿڱ��          HIS������
    '���Σ�Lng�ӿڱ�� ����    �Ժ��ֽ��㿨����
    '         strBlanceInfor  ����    ��||�ָ�: �ӿڱ��||���ѿ�ID(�ɴ�'')||���㷽ʽ||������||����||������ˮ��||����ʱ��(yyyy-mm-dd hh24:mi:ss)||��ע
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-06-18 11:33:22
    '˵������Ԥ�������������Ԥ��ʱ�����ȷ����ťʱ����(����ǰ����)
    '------------------------------------------------------------------------------------------------------------------------
    zlSimulationBrushCard = True
End Function
Public Function zlBrushCardToDeposit(ByVal lngԤ��ID As Long, ByVal lng�ӿڱ�� As Long, ByRef strBlanceInfor As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ����Ԥ������
    '��Σ� lngԤ��ID-Ԥ��ID
    '           lng�ӿڱ��-���㿨���
    '���Σ�strBlanceInfor-����ˢ����Ϣ:
    '         ��||�ָ�: �ӿڱ��||���ѿ�ID(�ɴ�'')||���㷽ʽ||������||����||������ˮ��||����ʱ��(yyyy-mm-dd hh24:mi:ss)||��ע
    '���أ��ɹ�����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-18 11:33:22
    '˵������Ԥ�������������Ԥ��ʱ�����ȷ����ťʱ����(�����е���)
    '          ����һ��Ҫ������ȷ,�������ֳ������
    '------------------------------------------------------------------------------------------------------------------------
    zlBrushCardToDeposit = True
End Function
 

Public Function zlDepositDel(ByVal lngԤ��ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���Ԥ������
    '��Σ� lngԤ��ID-Ԥ��ID
    '���أ��ɹ�����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-18 11:33:22
    '˵����1. ��Ԥ��������е��˿�, ��ȷ��ʱ(�����е���)�����øý���
     '------------------------------------------------------------------------------------------------------------------------
     zlDepositDel = True
End Function


Public Function zlReadCard(frmMain As Object, ByRef strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ӿ�:�ݲ����ǲ�����ؽӿ�
    '���:frmMain-���õ�������
    '����:strCardNo-���صĿ���
    '����:��������    True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:44:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject, strPath As String, strFile As String, strCurCardNo As String
    Dim objText As TextStream
     
    Err = 0: On Error GoTo Errhand:
    strPath = App.Path & "\����"
    If objFile.FolderExists(strPath) = False Then objFile.CreateFolder (strPath)
    strFile = strPath & "\CardInfor.txt"
    If objFile.FileExists(strFile) = False Then objFile.CreateTextFile strFile, True
    
    With gTy_TestBug
        '������
        If .objSquareCard Is Nothing Then Exit Function
        If .bytType = 1 Then
            If .strStartNo <> "" Then
                '������ǽ�����
                '����̫��,���������10���ڵĵ���
                strCurCardNo = Val(Replace(.strStartNo, .objSquareCard.ǰ׺�ı�, "")) + Int((Rnd * 100))
                strCurCardNo = String(.objSquareCard.���ų��� - Len(strCurCardNo), "0") & strCurCardNo
                strCardNo = .objSquareCard.ǰ׺�ı� & strCurCardNo
            Else
                '�������һ������
                strCurCardNo = Int(Rnd * IIf(.objSquareCard.���ų��� > 5, 10000, .objSquareCard.���ų���))
                strCurCardNo = String(.objSquareCard.���ų��� - Len(strCurCardNo), "0") & strCurCardNo
                strCardNo = .objSquareCard.ǰ׺�ı� & strCurCardNo
            End If
            '������ı��м��������
            Set objText = objFile.OpenTextFile(strFile, ForAppending)
            objText.WriteLine "1" & vbTab & strCardNo: objText.Close
        Else
            '��ȡ�Ѿ�ˢ����������
            strCardNo = Readģ�⿨��(strFile)
        End If
    End With
    zlReadCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Public Function zlCloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    '����ģ�壬��ʵ�־��幦��
End Function
Private Sub Class_Terminate()
    Call zlCloseCard
End Sub

Public Function zlReleaseComponent() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ͷ���Դ�ӿ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-06-24 09:44:54
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlReleaseComponent = True
End Function
