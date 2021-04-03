VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_URISCAN_PRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:URISCAN ��Һ����
'����:Sysmex
'����:����
'�û�:��������Ժ

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = Chr(6)
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    Dim varline As Variant, strItem As String, strItem1 As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    Const strStartChr As String = ""
    Const strEndChr As String = ""
    Dim strNowData As String
    Dim strData As String
    Dim intLoop As Integer
    
    
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    
    'û���ҵ���ʼ���ͽ�����ʱ�˳�
    If InStr(strSource, strStartChr) = 0 Or InStr(strSource, strEndChr) = 0 Then
        strReserved = strSource
    End If
    
    'ȡ�õ�ǰһ�Σ�ʣ�µķ��ء��´��ٽ�����
    strNowData = Mid(strSource, InStr(strSource, strStartChr), InStr(strSource, strEndChr))
    strReserved = Mid(strSource, InStr(strSource, strEndChr) + 1)
    
    '��ʼ����
    varline = Split(strNowData, vbCrLf)
    
    strData = Mid(varline(0), InStr(varline(0), ":") + 1) & "|"
    strData = strData & Val(Mid(varline(1), InStr(varline(1), ":") + 1)) & "| |��Һ|0"
    
    'ָ��
    For intLoop = LBound(varline) To UBound(varline)
        
        If InStr(",BLD,BIL,URO,KET,PRO,NIT,GLU,p.H,S.G,LEU,VTC,", "," & Mid(varline(intLoop), 1, 3) & ",") > 0 Then
            strItem = Trim(Mid(varline(intLoop), 4))
            strItem = Replace(strItem, "RBC/ul", "")
            If strItem <> "" Then strData = strData & "|" & Mid(varline(intLoop), 1, 3) & "|" & strItem
        End If
    Next
    strResult = strData
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:������Ҫ���������͵ı걾��Ϣ
    '           �걾֮����||�ָ�
    '           Ԫ��֮����|�ָ�
    '           ��0��Ԫ�أ�����ʱ��
    '           ��1��Ԫ�أ��������
    '           ��2��Ԫ�أ�������
    '           ��3��Ԫ�أ��걾
    '           ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '           ��5��9Ԫ�أ�ϵͳ����
    '           �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    '   iSendStep  : ����
    '   blnSuccess : �Ƿ���ʧ��
    '   strResponse: ������Ӧ����Ϣ
    '   blnUndo    : �Ƿ�ȡ��.
    '
    '   ����ֵ�������������ִ���ָ�
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function



