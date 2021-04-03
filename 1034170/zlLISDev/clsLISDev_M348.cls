VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_M348"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:U348 Ѫ��������
'����:
'����:����
'�û�:��������Ժ
Private mstrCode As String

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
    
    Dim varLine As Variant, strItem As String, strItem1 As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    Const strStartChr As String = "---- Chiron Diagnostics 348 ----"
    Const strEndChr As String = ""
    Dim strNowData As String
    Dim strData As String
    Dim intLoop As Integer
    Dim strLineData As String
    Dim aItem As Variant
    Dim lngS As Long, lngE As Long
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    
    mstrCode = mstrCode & strSource
    lngS = InStr(mstrCode, strStartChr)
    lngE = 0
    If lngS > 0 Then
        lngE = InStr(lngS, mstrCode, strEndChr)
    End If
    'û���ҵ���ʼ���ͽ�����ʱ�˳�
    If lngS <= 0 Or lngE <= 0 Then Exit Sub
    
    
    'ȡ�õ�ǰһ�Σ�ʣ�µķ��ء��´��ٽ�����
    strNowData = Mid(mstrCode, lngS, lngE - lngS + 1)
    mstrCode = Mid(mstrCode, lngE + 1)
    
    '��ʼ����
    varLine = Split(strNowData, vbCrLf)
    '����
    strLineData = DelSapce(CStr(varLine(2)))
    aItem = Split(strLineData, Space(1))
    strData = Format(aItem(4) & "-" & TrunMonth(aItem(3)) & "-" & aItem(2) & " " & aItem(1), "yyyy-mm-dd hh:mm:ss")
    
    '�걾��
    strLineData = DelSapce(CStr(varLine(5)))
    aItem = Split(strLineData, Space(1))
    strData = strData & "|" & aItem(2) & "| |Ѫ��|0"
    
    'ָ��
    For intLoop = 7 To UBound(varLine)
        strLineData = DelSapce(CStr(varLine(intLoop)))
        aItem = Split(strLineData, Space(1))
        If UBound(aItem) >= 3 Then
            If InStr(strData, "|" & aItem(1) & "|") <= 0 And aItem(1) <> "=" Then
                strData = strData & "|" & Trim(aItem(1)) & "|" & Trim(aItem(2))
            End If
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

Private Function DelSapce(strLine As String) As String
    '����       ɾ������Ŀո�
    Dim intLoop  As Integer
    Dim strNow As String
    strNow = strLine
    For intLoop = 20 To 0 Step -1
        strNow = Replace(strNow, Space(intLoop), Space(1))
    Next
    DelSapce = strNow
End Function
Private Function TrunMonth(ByVal strMonth As String) As String
    'ת�����ĵ��·�Ϊ����
    Select Case UCase(strMonth)
        Case "һ��": TrunMonth = "01"
        Case "����": TrunMonth = "02"
        Case "����": TrunMonth = "03"
        Case "����": TrunMonth = "04"
        Case "����": TrunMonth = "05"
        Case "����": TrunMonth = "06"
        Case "����": TrunMonth = "07"
        Case "����": TrunMonth = "08"
        Case "����": TrunMonth = "09"
        Case "ʮ��": TrunMonth = "10"
        Case "ʮһ": TrunMonth = "11"
        Case "ʮ����": TrunMonth = "12"
        
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select
End Function

