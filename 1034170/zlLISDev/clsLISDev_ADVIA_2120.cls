VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ADVIA_2120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ADVIA 2120ȫ�Զ�Ѫϸ��������
'����:�ݶ�
'����:���
'�û�:��ҽ��Ժ


Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    Dim strCheckBit As String
   
    GetStartCmd = Chr(&H30) & Chr(&H49) & Space(1) & Chr(&H7C) & "001" & Chr(&H7C) & Chr(&H2F) & Chr(&H78) & Chr(&H2F) & Chr(&H7C) & Chr(&HD) & Chr(&HA)
    strCheckBit = GetCheckSum(GetStartCmd)
    GetStartCmd = Chr(&H2) & GetStartCmd & Chr("&H" & strCheckBit) & Chr(&H3)
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
        
    '�豸��׼�ӿ�-----U200
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������^�Ƿ���^�����  ����ֻ���������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����

    Dim strResponseFirst As String
    Dim strResponseThird As String
    Dim strCheckBit As String
    Dim strResponse As String
    Dim strMsgTag As String
    Dim strTmp As String, varTmp As Variant
    Dim row As Integer, Rows As Integer
    
    Dim str_�걾�� As String, str_�������� As String, str_�걾���� As String
    Dim str_������Ŀ As String
    
    If strSource <> Chr(&H30) And strSource <> Chr(&H15) Then
        If Not (InStr(strSource, Chr(&H2)) > 0 And InStr(strSource, Chr(&H3)) > 0 _
            And InStr(strSource, Chr(&H3)) > InStr(strSource, Chr(&H2))) Then
            'If strSource Like "?" & Chr(&H2) & "*" & Chr(&H3) = False Then
            strReserved = strSource
            Exit Sub
        End If
    End If
    
    If Trim(strSource) = Chr(&H30) Then
        '����0ʱ���͹̶��ַ���
        strCmd = Chr(&H31) & Chr(&H53) & Space(5) & Chr(&HD) & Chr(&HA)
        strCheckBit = GetCheckSum(strCmd)
        strCmd = Chr(&H2) & strCmd & Chr("&H" & strCheckBit) & Chr(&H3)
    ElseIf Asc(Mid(strSource, 1, 1)) >= 49 And Asc(Mid(strSource, 1, 1)) <= 82 And strSource Like "?*" Then
        strResponseThird = Hex(Asc(Mid(strSource, 3, 1)) + 1)
        strResponseFirst = Hex(Asc(Mid(strSource, 3, 1)))
        strMsgTag = Mid(strSource, 4, 1)
        If strMsgTag = "S" Then
            strCmd = Chr("&H" & strResponseThird) & Chr(&H53) & Space(5) & Chr(&HD) & Chr(&HA)
            strCheckBit = GetCheckSum(strCmd)
            strCmd = Chr("&H" & strResponseFirst) & Chr(&H2) & strCmd & Chr("&H" & strCheckBit) & Chr(&H3)
        ElseIf strMsgTag = "R" Then
            str_�걾�� = Val(Mid(strSource, 6, 14))
            
            '04/11/10 17:23:23
            str_�������� = Mid(strSource, 44, 2) & "-" & Mid(strSource, 38, 2) & "-" & Mid(strSource, 41, 2)
            str_�������� = Format(str_��������, "yyyy-MM-dd")
            
            varTmp = Split(strSource, Chr(&HD) & Chr(&HA))
            strTmp = varTmp(1): Rows = Len(strTmp) / 9
            
            For row = 1 To Rows
                str_������Ŀ = str_������Ŀ & "|" & Trim(Mid(strTmp, 1, 3)) & "|" & Trim(Mid(strTmp, 4, 6))
                strTmp = Mid(strTmp, 10)
            Next
        
            strCmd = Chr("&H" & strResponseThird) & Chr(&H5A) & Space(18) & Chr(&H30) & Chr(&HD) & Chr(&HA)
            strCheckBit = GetCheckSum(strCmd)
            strCmd = Chr("&H" & strResponseFirst) & Chr(&H2) & strCmd & Chr("&H" & strCheckBit) & Chr(&H3)
        End If

    Else
        strCmd = Chr(&H30) & Chr(&H49) & Space(1) & Chr(&H7C) & "001" & Chr(&H7C) & Chr(&H2F) & Chr(&H78) & Chr(&H2F) & Chr(&H7C) & Chr(&HD) & Chr(&HA)
        strCheckBit = GetCheckSum(strCmd)
        strCmd = "  " & Chr(&H2) & strCmd & Chr("&H" & strCheckBit) & Chr(&H3)
    End If
    
    If Val(str_�걾��) <> 0 And str_������Ŀ <> "" Then
        strResult = str_�������� & "|" & str_�걾�� & "| |ȫѪ|0" & str_������Ŀ
    End If
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '   strSample:������Ҫ���������͵ı걾��Ϣ
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
    SendSample = ""
End Function

Private Function GetCheckSum(ByVal strResponse As String) As String
    Dim i As Integer
    Dim lngLength As Long
    Dim lngResult As Long

    lngLength = Len(strResponse)
    lngResult = Asc(Left(strResponse, 1))
    
    For i = 2 To lngLength
        lngResult = lngResult Xor Asc(Mid(strResponse, i, 1))
    Next
    
    If lngResult = 3 Then
        lngResult = 127
    End If
    
    GetCheckSum = Hex(lngResult)
End Function

