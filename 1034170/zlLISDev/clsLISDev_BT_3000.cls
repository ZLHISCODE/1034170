VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BT_3000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:BT-3000
'����:�����
'�û�:������������ҽԺ
'����:����
Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
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
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    Dim iStrPoint As Integer, iStrPointBK As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean
    Dim intQC As Integer       '�ʿ�λ��
    Dim intDataType As Integer '�������� 0�������1������
    Dim intEnd As Integer      '���λ��
    Dim strStartChr As String, strEndChr As String
    Dim strNowData As String
    Dim strTmp As String
    Dim blnEmergency As String
    Dim str�걾�� As String
    Dim strItem As String
    
    On Error GoTo errHandle
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    strStartChr = Chr(&O2)
    strEndChr = Chr(&O3)
    
    'û�п�ʼ�ͽ���ʱ�˳�
    If InStr(strSource, strStartChr) <= 0 Then strReserved = strSource
    If InStr(strSource, strEndChr) <= 0 Then strReserved = strSource
    
    strNowData = Mid(strSource, InStr(strSource, strStartChr), InStr(strSource, strEndChr) + Len(strEndChr))
    strReserved = Mid(strSource, InStr(strSource, strEndChr) + Len(strEndChr))
    
    strNowData = Mid(strNowData, InStr(strNowData, Chr(&H54)) + 1)
    strTmp = Mid(strNowData, 1, 10)
    blnEmergency = IIf(InStr(strTmp, "Batch") > 0, True, False)
    str�걾�� = Replace(Replace(strTmp, "Batch", ""), " ", "")
    
    strData = Format(Now, "yyyy-mm-dd hh:mm:ss") & "|" & str�걾�� & "| |Ѫ��|0"
    strTmp = Mid(strNowData, 12)
    Do While strTmp <> ""
        strItem = Mid(strTmp, 1, 21)
        If Len(strItem) >= 21 Then
            strTmp = Mid(strTmp, 22)
            strData = strData & "|" & Trim(Mid(strItem, 1, 3)) & "|" & Val(Mid(strItem, 16))
        Else
            strTmp = ""
        End If
    Loop
    
    If strData <> "" Then strResult = strData
    Exit Sub
errHandle:
    Call WriteLog("BT_3000 analyse", "�������" & Err.Description, "")
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    
    
    
End Function

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    'ָ�붨λ����һ������֮��
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = Chr(10) Then
            iPoint = iPoint + 1
            Exit Do
        End If
        iPoint = iPoint + 1
    Loop
End Sub

Private Function ReadSubString(ByVal strSource As String, iPoint As Integer, Optional ByVal Length As Integer = 0) As String
    '��ȡ�Ӵ���ͬʱָ���ƶ�����ȡ���ַ���֮��
    On Error Resume Next
    If Length = 0 Then
        ReadSubString = Mid(strSource, iPoint)
        iPoint = Len(strSource) + 1
    Else
        ReadSubString = Mid(strSource, iPoint, Length)
        iPoint = iPoint + Length
    End If
End Function



