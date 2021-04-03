VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MD550"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:MD-550
'����:����ҽ���豸���޹�˾
'����:�¶�
'�û�:����ҽԺ

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
    
    Dim varLine As Variant, varItem As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String, lngLength As Long, strRecord As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    If strSource Like "*[" & Chr(&H23&) & "]" & Chr(&H24&) & "*" & Chr(&H24&) & Chr(&H26&) & "*" Or _
       strSource Like "*" & Chr(&H27&) & Chr(&H28&) & "*" & Chr(&H28&) & Chr(&H2A&) & "*" Then
       
        strSource = Replace(strSource, Chr(&H23&) & Chr(&H24&), Chr(&H27&) & Chr(&H28&))
        strSource = Replace(strSource, Chr(&H24&) & Chr(&H26&), Chr(&H28&) & Chr(&H2A&))
        
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
            lngLength = InStr(strReserved, Chr(&H2A&)) - InStr(strReserved, Chr(&H27&))
            strLine = Mid(strReserved, InStr(strReserved, Chr(&H27&)), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, Chr(&H2A&)) + 1)
            
            str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = "": strTmp = ""
            '����
            strItem = ""
            strLine = Replace(strLine, Chr(&H27&) & Chr(&H28&), "")
            strLine = Replace(strLine, Chr(&H28&) & Chr(&H2A&), "")
            Do While Len(strLine) >= 64
                strRecord = Mid(strLine, 1, 64)
                strLine = Mid(strLine, 65)
                str�걾�� = Val(Mid(strRecord, 1, 3)) & IIf(Mid(strRecord, 5, 1) = "1", "^1", "")
                str�����˱걾�ʿ� = " |" & Switch(Mid(strRecord, 4, 1) = "0", "ѪҺ", Mid(strRecord, 4, 1) = "1", "��Һ", Mid(strRecord, 4, 1) = "2", "����") & "|0"
                str���� = Mid(strRecord, 7, 4) & "-" & Mid(strRecord, 11, 2) & "-" & Mid(strRecord, 13, 2)
                
                strItem = IIf(InStr(Mid(strRecord, 22, 6), ".") > 0, "|01|" & Trim(Mid(strRecord, 22, 6)), "") & _
                          IIf(InStr(Mid(strRecord, 28, 6), ".") > 0, "|02|" & Trim(Mid(strRecord, 28, 6)), "") & _
                          IIf(InStr(Mid(strRecord, 34, 6), ".") > 0, "|03|" & Trim(Mid(strRecord, 34, 6)), "") & _
                          IIf(InStr(Mid(strRecord, 40, 6), ".") > 0, "|04|" & Trim(Mid(strRecord, 40, 6)), "") & _
                          IIf(InStr(Mid(strRecord, 46, 6), ".") > 0, "|05|" & Trim(Mid(strRecord, 46, 6)), "") & _
                          IIf(InStr(Mid(strRecord, 52, 6), ".") > 0, "|06|" & Trim(Mid(strRecord, 52, 6)), "") & _
                          IIf(InStr(Mid(strRecord, 58, 6), ".") > 0, "|07|" & Trim(Mid(strRecord, 58, 6)), "")
                If strItem <> "" And Val(str�걾��) <> 0 Then
                    strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
                End If
            Loop
            If Not strReserved Like "*" & Chr(&H27&) & Chr(&H28&) & "*" & Chr(&H28&) & Chr(&H2A&) & "*" Then
                blnEnd = True
            End If
        Loop
    Else
        '�յ�����Ϣ��ȫ,����,���´��ٽ���
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    'strCmd = Chr(6)
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
